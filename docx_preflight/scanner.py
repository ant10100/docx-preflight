"""Bounded package reader and OOXML heuristics. Never extract or follow links."""

from collections import Counter
from pathlib import Path
import io
import re
import stat
import struct
import xml.etree.ElementTree as ET
from xml.parsers import expat
import zipfile
import zlib

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
STRICT_W = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
MAX_ARCHIVE = 32 * 1024 * 1024
MAX_PART = 8 * 1024 * 1024
MAX_TOTAL = 64 * 1024 * 1024
MAX_ENTRIES = 256
MAX_RATIO = 200
MAX_DEPTH = 100
MAX_NODES = 100000
MAX_NUMBERING_STEPS = 200000
MAX_DIRECTORY = 256 * 1024
MAX_TOTAL_XML_NODES = 200000
MAX_ATTRIBUTES_PER_NODE = 256
MAX_TOTAL_XML_ATTRIBUTES = 200000
MAX_TOTAL_XML_BYTES = 16 * 1024 * 1024
MAX_NAMESPACE_URI_BYTES = 2048
MAX_XML_NAME_BYTES = 512
MAX_EXPANDED_XML_NAME_BYTES = 64 * 1024 * 1024


class Rejected(ValueError):
    """Only fixed, public diagnostic codes may be included in this exception."""


def _xml(data, remaining_nodes=None, remaining_attributes=None, remaining_names=None):
    # Expat recognizes declarations in UTF-16 too. A raw byte regex would not.
    parser = expat.ParserCreate()
    depth = nodes = 0
    namespaces = {'xml': 'http://www.w3.org/XML/1998/namespace'}
    namespace_frames = []
    missing_namespace = object()
    expanded_name_bytes = 0

    def bounded_utf8(value, maximum, code):
        if len(value) > maximum or len(value.encode('utf-8')) > maximum:
            raise Rejected(code)

    def account_name_size(characters):
        nonlocal expanded_name_bytes
        # Conservative Unicode storage estimate: four bytes per character,
        # room for Clark-notation braces/terminator and an object header.
        amount = 4 * (characters + 2) + 80
        expanded_name_bytes += amount
        if expanded_name_bytes > MAX_EXPANDED_XML_NAME_BYTES:
            raise Rejected('xml_expanded_name_limit')
        if remaining_names is not None:
            remaining_names[0] -= amount
            if remaining_names[0] < 0:
                raise Rejected('xml_expanded_name_limit')

    def account_expanded_name(name, attribute=False):
        prefix, separator, local = name.partition(':')
        if separator:
            if not prefix or not local or ':' in local or prefix not in namespaces:
                raise Rejected('xml_namespace_invalid')
            uri = namespaces[prefix]
        else:
            local = name
            uri = '' if attribute else namespaces.get('', '')
        # Account for the later Strict-to-Transitional normalization as well.
        uri_length = len(W) if uri == STRICT_W else len(uri)
        account_name_size(uri_length + len(local))

    def reject(*args):
        raise Rejected('xml_declaration_forbidden')

    def start(name, attributes):
        nonlocal depth, nodes
        depth += 1
        nodes += 1
        if len(attributes) > MAX_ATTRIBUTES_PER_NODE:
            raise Rejected('xml_attribute_limit')
        if remaining_attributes is not None:
            remaining_attributes[0] -= len(attributes)
            if remaining_attributes[0] < 0:
                raise Rejected('total_xml_complexity_limit')
        if remaining_nodes is not None:
            remaining_nodes[0] -= 1
            if remaining_nodes[0] < 0:
                raise Rejected('total_xml_complexity_limit')
        if depth > MAX_DEPTH or nodes > MAX_NODES:
            raise Rejected('xml_complexity_limit')
        # Do not enable Expat namespace expansion: that could itself allocate
        # amplified names before callbacks run. Track bounded bindings without
        # copying the whole scope map at each depth, then estimate ET's names.
        bounded_utf8(name, MAX_XML_NAME_BYTES, 'xml_name_limit')
        changes = []
        for attribute, value in attributes.items():
            bounded_utf8(attribute, MAX_XML_NAME_BYTES, 'xml_name_limit')
            if attribute == 'xmlns' or attribute.startswith('xmlns:'):
                bounded_utf8(value, MAX_NAMESPACE_URI_BYTES, 'xml_namespace_uri_limit')
                prefix = '' if attribute == 'xmlns' else attribute[6:]
                if ':' in prefix or (attribute != 'xmlns' and not prefix):
                    raise Rejected('xml_namespace_invalid')
                account_name_size(len(attribute) + len(value))
                changes.append((prefix, namespaces.get(prefix, missing_namespace)))
                namespaces[prefix] = value
        namespace_frames.append(changes)
        account_expanded_name(name)
        for attribute in attributes:
            if attribute != 'xmlns' and not attribute.startswith('xmlns:'):
                account_expanded_name(attribute, attribute=True)

    def end(*args):
        nonlocal depth
        for prefix, previous in reversed(namespace_frames.pop()):
            if previous is missing_namespace:
                del namespaces[prefix]
            else:
                namespaces[prefix] = previous
        depth -= 1

    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    root = ET.fromstring(data)
    for node in root.iter():
        if node.tag.startswith('{' + STRICT_W + '}'):
            node.tag = node.tag.replace(STRICT_W, W, 1)
        node.attrib = {k.replace(STRICT_W, W, 1): v for k, v in node.attrib.items()}
    return root


def _check_extra(extra):
    position = 0
    while position < len(extra):
        if len(extra) - position < 4:
            raise Rejected('invalid_central_directory')
        kind, length = struct.unpack_from('<HH', extra, position)
        position += 4
        if length > len(extra) - position:
            raise Rejected('invalid_central_directory')
        if kind == 1:
            raise Rejected('unsupported_zip64')
        position += length


def _check_directory(source, size):
    """Validate bounded ZIP32 directory records before ZipFile creates ZipInfo.

    Only canonical single-disk ZIP32 offsets are accepted. Neither declared
    counts nor declared record lengths may bypass the actual record walk.
    """
    tail_size = min(size, 22 + 65535)
    source.seek(size - tail_size)
    tail = source.read(tail_size)
    position = tail.rfind(b'PK\x05\x06')
    if position < 0 or position + 22 > len(tail):
        raise Rejected('invalid_central_directory')
    end = struct.unpack_from('<4s4H2IH', tail, position)
    _, disk, start_disk, disk_count, declared_count, directory_size, offset, comment_size = end
    if position + 22 + comment_size != len(tail):
        raise Rejected('invalid_central_directory')
    absolute_end = size - tail_size + position
    if (disk_count == 65535 or declared_count == 65535 or
            directory_size == 0xffffffff or offset == 0xffffffff or
            (position >= 20 and tail[position - 20:position - 16] == b'PK\x06\x07')):
        raise Rejected('unsupported_zip64')
    if disk != 0 or start_disk != 0 or disk_count != declared_count:
        raise Rejected('unsupported_multidisk_or_count')
    if declared_count > MAX_ENTRIES:
        raise Rejected('entry_count_limit')
    if directory_size > MAX_DIRECTORY:
        raise Rejected('central_directory_size_limit')
    if offset + directory_size != absolute_end or offset > absolute_end:
        raise Rejected('invalid_central_directory')
    source.seek(offset)
    count = 0
    consumed = 0
    while consumed < directory_size:
        count += 1
        if count > MAX_ENTRIES:
            raise Rejected('entry_count_limit')
        if directory_size - consumed < 46:
            raise Rejected('invalid_central_directory')
        header = source.read(46)
        if len(header) != 46 or header[:4] != b'PK\x01\x02':
            raise Rejected('invalid_central_directory')
        fields = struct.unpack('<4s6H3I5H2I', header)
        needed, compressed, expanded = fields[2], fields[8], fields[9]
        name_size, extra_size, file_comment, member_disk, local_offset = fields[10:14] + (fields[16],)
        if needed >= 45 or compressed == 0xffffffff or expanded == 0xffffffff or local_offset == 0xffffffff:
            raise Rejected('unsupported_zip64_or_version')
        if member_disk:
            raise Rejected('unsupported_multidisk_or_count')
        if not name_size or name_size > 960 or local_offset + 30 > offset:
            raise Rejected('invalid_central_directory')
        record_size = 46 + name_size + extra_size + file_comment
        if record_size > directory_size - consumed:
            raise Rejected('invalid_central_directory')
        name = source.read(name_size)
        extra = source.read(extra_size)
        if len(name) != name_size or len(extra) != extra_size:
            raise Rejected('invalid_central_directory')
        _check_extra(extra)
        source.seek(file_comment, 1)
        consumed += record_size
        next_record = source.tell()
        source.seek(local_offset)
        local_header = source.read(30)
        if len(local_header) != 30 or local_header[:4] != b'PK\x03\x04':
            raise Rejected('invalid_local_header')
        local = struct.unpack('<4s5H3I2H', local_header)
        if local[1] >= 45 or local[7] == 0xffffffff or local[8] == 0xffffffff:
            raise Rejected('unsupported_zip64_or_version')
        if (local[2] != fields[3] or local[3] != fields[4] or local[9] != name_size
                or local_offset + 30 + local[9] + local[10] + compressed > offset):
            raise Rejected('invalid_local_header')
        if not local[2] & 8 and (local[6], local[7], local[8]) != (fields[7], compressed, expanded):
            raise Rejected('invalid_local_header')
        if source.read(local[9]) != name:
            raise Rejected('invalid_local_header')
        local_extra = source.read(local[10])
        if len(local_extra) != local[10]:
            raise Rejected('invalid_local_header')
        _check_extra(local_extra)
        source.seek(next_record)
    if count != declared_count:
        raise Rejected('inconsistent_directory_count')
    source.seek(0)


def _read(path):
    # Reject huge directories before taking a bounded immutable snapshot.
    # Recheck that snapshot so concurrent file edits cannot make ZipFile parse
    # a different directory from the one that passed the preflight.
    with Path(path).open('rb') as physical:
        size = physical.seek(0, 2)
        if size > MAX_ARCHIVE:
            raise Rejected('archive_size_limit')
        _check_directory(physical, size)
        data = physical.read(size + 1)
        if len(data) > MAX_ARCHIVE or len(data) != size:
            raise Rejected('input_changed_or_size_limit')
    with io.BytesIO(data) as source:
        _check_directory(source, len(data))
        with zipfile.ZipFile(source) as archive:
            members = archive.infolist()
            if len(members) > MAX_ENTRIES:
                raise Rejected('entry_count_limit')
            seen = set()
            total = 0
            for entry in members:
                # ZipInfo normalizes backslashes on Windows; inspect original spelling.
                name = entry.orig_filename
                segments = name.split('/')
                if (not name or '\\' in name or ':' in name or name.startswith('/')
                        or any(s in ('.', '..', '') for s in segments[:-1])
                        or '..' in segments or len(name) > 240 or '\x00' in name):
                    raise Rejected('unsafe_member_name')
                if name.casefold() in seen:
                    raise Rejected('duplicate_member')
                seen.add(name.casefold())
                if entry.flag_bits & 1:
                    raise Rejected('encrypted_member')
                mode = entry.external_attr >> 16
                if stat.S_ISLNK(mode):
                    raise Rejected('symlink_member')
                if entry.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                    raise Rejected('unsupported_compression')
                total += entry.file_size
                if (entry.file_size > MAX_PART or total > MAX_TOTAL
                        or entry.file_size > MAX_RATIO * max(entry.compress_size, 1)):
                    raise Rejected('expanded_size_or_ratio_limit')
            roots = {}
            binaries = []
            actual_total = 0
            xml_total = 0
            remaining_nodes = [MAX_TOTAL_XML_NODES]
            remaining_attributes = [MAX_TOTAL_XML_ATTRIBUTES]
            remaining_names = [MAX_EXPANDED_XML_NAME_BYTES]
            for entry in members:
                if entry.is_dir():
                    continue
                with archive.open(entry) as stream:
                    data = stream.read(MAX_PART + 1)
                actual_total += len(data)
                if len(data) > MAX_PART or actual_total > MAX_TOTAL:
                    raise Rejected('expanded_size_or_ratio_limit')
                if entry.filename.lower().endswith(('.xml', '.rels')):
                    xml_total += len(data)
                    if xml_total > MAX_TOTAL_XML_BYTES:
                        raise Rejected('total_xml_size_limit')
                    roots[entry.filename] = _xml(data, remaining_nodes, remaining_attributes, remaining_names)
                else:
                    binaries.append(entry.filename)
            return roots, binaries


def _numbering(roots, add):
    # One XML walk builds first-child lookups, definitions and paragraph text
    # flags. No paragraph searches a shared style/num/abstractNum subtree.
    steps = 0

    def step():
        nonlocal steps
        steps += 1
        if steps > MAX_NUMBERING_STEPS:
            raise Rejected('numbering_complexity_limit')

    prefix = '{' + W + '}'
    indexed_tags = {prefix + tag for tag in (
        'styles', 'docDefaults', 'pPrDefault', 'style', 'p', 'pPr', 'numPr',
        'numbering', 'num', 'abstractNum', 'lvlOverride', 'lvl')}
    queried_tags = {prefix + tag for tag in (
        'docDefaults', 'pPrDefault', 'pPr', 'basedOn', 'pStyle', 'numPr',
        'numId', 'ilvl', 'abstractNumId', 'numStyleLink', 'lvl', 'numFmt')}
    first_children = {}
    node_values = {}
    canonical_values = {}
    valid_ids = {}
    valid_levels = {}
    styles_root = roots.get('word/styles.xml')
    numbering_root = roots.get('word/numbering.xml')
    styles = {}
    default_style = None
    nums = {}
    abstracts = {}
    paragraphs = []

    def canonical(value):
        if value is None:
            return None
        found = canonical_values.get(value)
        if found is None:
            canonical_values[value] = value
            return value
        return found

    def child(node, tag):
        return first_children.get(node, {}).get(prefix + tag)

    def value(node, tag):
        found = child(node, tag)
        return node_values.get(found)

    def numprops(properties):
        numbering = child(properties, 'numPr')
        return {key: found for key in ('numId', 'ilvl')
                if (found := value(numbering, key)) is not None}

    def walk(node, parent, name):
        nonlocal default_style
        step()  # Every element, including unmatched/ignored children, counts.
        if node.tag in queried_tags and prefix + 'val' in node.attrib:
            node_values[node] = canonical(node.get(prefix + 'val'))
        indexed = {} if node.tag in indexed_tags else None
        has_text = node.tag == prefix + 't' and bool((node.text or '').strip())
        levels = {}
        overrides = {}
        for element in node:
            descendant_text = walk(element, node, name)
            has_text = has_text or descendant_text
            if indexed is not None and element.tag in queried_tags:
                indexed.setdefault(element.tag, element)
            if node.tag == prefix + 'abstractNum' and element.tag == prefix + 'lvl':
                levels[canonical(element.get(prefix + 'ilvl'))] = value(element, 'numFmt')
            if node.tag == prefix + 'num' and element.tag == prefix + 'lvlOverride':
                level = child(element, 'lvl')
                # Distinguish an absent override level (fall back) from a
                # present level missing numFmt (unresolved). Last override wins.
                overrides[canonical(element.get(prefix + 'ilvl'))] = (level is not None, value(level, 'numFmt'))
        if indexed is not None:
            first_children[node] = indexed
        if styles_root is not None and parent is styles_root and node.tag == prefix + 'style' and node.get(prefix + 'type') == 'paragraph':
            style_id = canonical(node.get(prefix + 'styleId'))
            styles[style_id] = (numprops(child(node, 'pPr')), value(node, 'basedOn'))
            if node.get(prefix + 'default') in ('1', 'true', 'on'):
                default_style = style_id
        if numbering_root is not None and parent is numbering_root:
            if node.tag == prefix + 'num':
                nums[canonical(node.get(prefix + 'numId'))] = (value(node, 'abstractNumId'), overrides)
            elif node.tag == prefix + 'abstractNum':
                abstracts[canonical(node.get(prefix + 'abstractNumId'))] = (child(node, 'numStyleLink') is not None, levels)
        if node.tag == prefix + 'p' and name not in ('word/styles.xml', 'word/numbering.xml'):
            paragraphs.append((node, has_text))
        return has_text

    for name, root in roots.items():
        if name.startswith('word/'):
            walk(root, None, name)
    defaults = numprops(child(child(child(styles_root, 'docDefaults'), 'pPrDefault'), 'pPr'))
    for paragraph, has_text in paragraphs:
        step()
        properties = child(paragraph, 'pPr')
        style_id = value(properties, 'pStyle') or default_style
        chain = []
        visited = set()
        while style_id:
            step()
            if style_id in visited or style_id not in styles:
                add('unresolved_style', 'warning')
                break
            visited.add(style_id)
            props, style_id = styles[style_id]
            chain.append(props)
        effective = dict(defaults)
        for props in reversed(chain):
            effective.update(props)
        direct = numprops(properties)
        effective.update(direct)
        number_id = effective.get('numId')
        level = effective.get('ilvl', '0')
        if number_id in (None, '0'):
            continue
        # Shared, potentially long attribute strings are canonicalized once.
        # Classification is cached too, avoiding paragraphs times string length.
        if number_id not in valid_ids:
            valid_ids[number_id] = number_id.isdecimal()
        if not valid_ids[number_id]:
            add('unresolved_numbering', 'warning')
            continue
        if level not in valid_levels:
            valid_levels[level] = level.isdecimal() and 0 <= int(level) <= 8
        if not valid_levels[level]:
            add('unresolved_numbering', 'warning')
            continue
        if (defaults.get('numId') == number_id or
                (default_style in visited and not direct.get('numId'))):
            add('default_numbering', 'warning')
        abstract_id, overrides = nums.get(number_id, (None, {}))
        linked, levels = abstracts.get(abstract_id, (False, {}))
        if linked:
            add('unsupported_numbering_link', 'warning')
            continue
        present, fmt = overrides.get(level, (False, None))
        if not present:
            fmt = levels.get(level)
        if fmt is None:
            add('unresolved_numbering', 'warning')
        elif fmt == 'none':
            add('suppressed_numbering', 'info')
        else:
            add('list_paragraph', 'info')
            add('bullet_list' if fmt == 'bullet' else 'numbered_list', 'info')
            if not has_text:
                add('empty_list_paragraph', 'warning')
    return steps


def inspect_docx(path):
    """Return only fixed finding codes/counts. No paths, text, or metadata values.

    `ok` means no warning from implemented checks, never complete privacy safety.
    Inspection errors return `error` and must not be treated as a clean result.
    """
    counts = Counter()

    def add(code, severity='warning', count=1):
        counts[(severity, code)] += count

    report = {'schema_version': 1, 'status': 'ok', 'findings': [],
              'scope': 'Heuristic OOXML checks; not a privacy or rendering guarantee.'}
    try:
        roots, binaries = _read(path)
        document = roots.get('word/document.xml')
        content = roots.get('[Content_Types].xml')
        relationships = roots.get('_rels/.rels')
        if (document is None or document.tag != '{' + W + '}document'
                or content is None or content.tag != '{' + CT + '}Types'
                or relationships is None or relationships.tag != '{' + R + '}Relationships'):
            raise Rejected('invalid_docx_package')
        if not any(n.get('PartName') == '/word/document.xml' and n.get('ContentType') == MAIN for n in content):
            raise Rejected('unsupported_main_content_type')
        if not any(n.get('Type', '').endswith('/officeDocument') and
                   n.get('Target') in ('word/document.xml', '/word/document.xml') and
                   n.get('TargetMode', 'Internal') == 'Internal' for n in relationships):
            raise Rejected('unsupported_main_relationship')
        revision_names = {'ins', 'del', 'moveFrom', 'moveTo', 'moveFromRangeStart', 'moveToRangeStart', 'cellIns', 'cellDel', 'cellMerge', 'numberingChange'}
        for name, root in roots.items():
            lower = name.lower()
            if lower.startswith('docprops/'):
                for node in root:
                    if ''.join(node.itertext()).strip():
                        add('metadata')
            if lower.startswith('customxml/'):
                add('custom_xml')
            if re.search(r'/(comments[^/]*|people|person)\.xml$', lower):
                add('comments' if 'comments' in lower else 'collaborator_data')
            for node in root.iter():
                local = node.tag.rsplit('}', 1)[-1]
                if node.tag.startswith('{' + W + '}'):
                    if local in revision_names or local.endswith('PrChange'):
                        add('tracked_changes')
                    if local in ('vanish', 'webHidden') and node.get('{' + W + '}val', '1') not in ('0', 'false', 'off'):
                        add('hidden_text')
                    if local in ('instrText', 'fldSimple'):
                        add('field_code')
                    if local in ('commentRangeStart', 'commentReference'):
                        add('comment_reference')
                if any(k.rsplit('}', 1)[-1] in ('author', 'initials', 'userId', 'email') and v for k, v in node.attrib.items()):
                    add('collaborator_identity')
                if lower.endswith('.rels') and local == 'Relationship' and node.get('TargetMode', '').lower() == 'external':
                    add('external_relationship')
        if binaries:
            add('uninspected_binary', count=len(binaries))
        _numbering(roots, add)
    except Rejected as error:
        counts.clear()
        add(str(error), 'error')
        report['status'] = 'error'
    except (OSError, ValueError, LookupError, MemoryError, OverflowError, zipfile.BadZipFile, NotImplementedError,
            RuntimeError, ET.ParseError, expat.ExpatError, zlib.error):
        counts.clear()
        add('unreadable_or_invalid_input', 'error')
        report['status'] = 'error'
    if report['status'] != 'error' and any(severity == 'warning' for severity, _ in counts):
        report['status'] = 'warnings'
    report['findings'] = [{'severity': severity, 'code': code, 'count': count}
                          for (severity, code), count in sorted(counts.items())]
    return report
