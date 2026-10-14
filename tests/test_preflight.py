import contextlib
import hashlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import struct
import subprocess
import sys
import xml.etree.ElementTree as ET
import warnings
import zipfile

from docx_preflight import inspect_docx
from docx_preflight.cli import main
from docx_preflight.scanner import _numbering
import docx_preflight.scanner as scanner

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'


def package(path, body='', extra=None, compression=zipfile.ZIP_DEFLATED):
    parts = {
        '[Content_Types].xml': f'<Types xmlns="{CT}"><Override PartName="/word/document.xml" ContentType="{MAIN}"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>',
    }
    parts.update(extra or {})
    with zipfile.ZipFile(path, 'w', compression) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return path


def numbering():
    return f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>'


def directory_flood(path, count):
    """A huge directory with a forged EOCD count, built with constant memory."""
    local = struct.pack('<4s5H3I2H', b'PK\x03\x04', 20, 0, 0, 0, 0, 0, 0, 0, 1, 0) + b'x'
    central = struct.pack('<4s6H3I5H2I', b'PK\x01\x02', 20, 20, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0) + b'x'
    with Path(path).open('wb') as stream:
        stream.write(local)
        for _ in range(count // 1000):
            stream.write(central * 1000)
        stream.write(central * (count % 1000))
        stream.write(struct.pack('<4s4H2IH', b'PK\x05\x06', 0, 0, 1, 1, len(central) * count, len(local), 0))


class PreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'synthetic.docx'

    def scan(self, body='', extra=None):
        package(self.path, body, extra)
        return inspect_docx(self.path)

    def codes(self, report):
        return {f['code'] for f in report['findings']}

    def test_plain_and_read_only(self):
        package(self.path, '<w:p><w:r><w:t>Synthetic paragraph</w:t></w:r></w:p>')
        before = hashlib.sha256(self.path.read_bytes()).digest()
        report = inspect_docx(self.path)
        self.assertEqual(report['status'], 'ok')
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())

    def test_valid_list_is_information(self):
        report = self.scan('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>', {'word/numbering.xml': numbering()})
        self.assertEqual(report['status'], 'ok')
        self.assertIn('list_paragraph', self.codes(report))

    def test_broken_numbering_warns(self):
        report = self.scan('<w:p><w:pPr><w:numPr><w:numId w:val="42"/></w:numPr></w:pPr></w:p>')
        self.assertIn('unresolved_numbering', self.codes(report))

    def test_style_inheritance_and_zero_disable(self):
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style></w:styles>'
        extra = {'word/styles.xml': styles, 'word/numbering.xml': numbering()}
        report = self.scan('<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>', extra)
        self.assertIn('list_paragraph', self.codes(report))
        report = self.scan('<w:p><w:pPr><w:pStyle w:val="Child"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>', extra)
        self.assertNotIn('list_paragraph', self.codes(report))

    def test_default_numbering_warns(self):
        styles = f'<w:styles xmlns:w="{W}"><w:docDefaults><w:pPrDefault><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:pPrDefault></w:docDefaults></w:styles>'
        report = self.scan('<w:p><w:r><w:t>Paragraph</w:t></w:r></w:p>', {'word/styles.xml': styles, 'word/numbering.xml': numbering()})
        self.assertIn('default_numbering', self.codes(report))

    def test_privacy_categories_no_values(self):
        extra = {
            'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>SYNTHETIC_PRIVATE_MARKER</dc:creator><cp:lastModifiedBy>SYNTHETIC_PRIVATE_MARKER</cp:lastModifiedBy></cp:coreProperties>',
            'word/comments.xml': f'<w:comments xmlns:w="{W}"><w:comment w:author="SYNTHETIC_PRIVATE_MARKER"/></w:comments>',
            'word/_rels/document.xml.rels': f'<Relationships xmlns="{R}"><Relationship TargetMode="External" Target="https://example.invalid/SYNTHETIC_PRIVATE_MARKER"/></Relationships>',
        }
        report = self.scan('<w:p><w:ins w:author="SYNTHETIC_PRIVATE_MARKER"/></w:p>', extra)
        self.assertTrue({'metadata', 'comments', 'tracked_changes', 'external_relationship'} <= self.codes(report))
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', json.dumps(report))
        self.assertNotIn(str(self.path), json.dumps(report))

    def test_additional_privacy_surfaces(self):
        report = self.scan('<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:instrText> INCLUDETEXT synthetic </w:instrText></w:r></w:p>', {'customXml/item1.xml': '<data/>', 'word/media/image1.png': b'synthetic'})
        self.assertTrue({'hidden_text', 'field_code', 'custom_xml', 'uninspected_binary'} <= self.codes(report))

    def test_strict_namespace(self):
        package(self.path, '<w:p><w:ins/></w:p>')
        with zipfile.ZipFile(self.path) as z:
            parts = {n: z.read(n).replace(W.encode(), b'http://purl.oclc.org/ooxml/wordprocessingml/main') for n in z.namelist()}
        package(self.path, extra=parts)
        self.assertIn('tracked_changes', self.codes(inspect_docx(self.path)))

    def test_bad_zip_and_missing_main(self):
        self.path.write_bytes(b'not a zip')
        self.assertEqual(inspect_docx(self.path)['status'], 'error')
        with zipfile.ZipFile(self.path, 'w') as z:
            z.writestr('unrelated.xml', '<x/>')
        self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_unsafe_names_duplicates_and_xml(self):
        for name in ['../escape.xml', '/absolute.xml']:
            report = self.scan(extra={name: '<x/>'})
            self.assertEqual(report['status'], 'error')
        package(self.path, extra={'word/bad.xml': '<x/>'})
        self.path.write_bytes(self.path.read_bytes().replace(b'word/bad.xml', b'word\\bad.xml'))
        self.assertEqual(inspect_docx(self.path)['status'], 'error')
        package(self.path)
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            with zipfile.ZipFile(self.path, 'a') as z:
                z.writestr('word/document.xml', '<x/>')
        self.assertEqual(inspect_docx(self.path)['status'], 'error')
        self.assertEqual(self.scan(extra={'word/comments.xml': '<broken>'})['status'], 'error')

    def test_dtd_utf8_and_utf16(self):
        evil = '<!DOCTYPE x [<!ENTITY v "synthetic">]><x>&v;</x>'
        for data in [evil.encode(), evil.encode('utf-16')]:
            self.assertEqual(self.scan(extra={'customXml/item1.xml': data})['status'], 'error')

    def test_limits(self):
        self.assertEqual(self.scan(extra={'word/media/big.bin': b'X' * 1000000})['status'], 'error')
        self.assertEqual(self.scan(extra={f'part{i}.xml': '<x/>' for i in range(257)})['status'], 'error')
        self.assertEqual(self.scan(extra={'customXml/item1.xml': '<x>' * 110 + '</x>' * 110})['status'], 'error')

    def test_encrypted_and_corrupt_crc(self):
        package(self.path)
        data = bytearray(self.path.read_bytes())
        central = data.index(b'PK\x01\x02')
        struct.pack_into('<H', data, 6, 1)
        struct.pack_into('<H', data, central + 8, 1)
        self.path.write_bytes(data)
        self.assertIn('encrypted_member', self.codes(inspect_docx(self.path)))
        package(self.path)
        data = bytearray(self.path.read_bytes())
        central = data.index(b'PK\x01\x02')
        struct.pack_into('<I', data, central + 16, 0)
        self.path.write_bytes(data)
        self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_archive_part_total_and_nodes_limits(self):
        for constant, limit in [('MAX_ARCHIVE', 10), ('MAX_PART', 10), ('MAX_TOTAL', 10), ('MAX_NODES', 1)]:
            package(self.path)
            with patch('docx_preflight.scanner.' + constant, limit):
                self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_unused_numbering_and_numbered_lists(self):
        self.assertNotIn('list_paragraph', self.codes(self.scan('<w:p/>', {'word/numbering.xml': numbering()})))
        report = self.scan('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>', {'word/numbering.xml': numbering().replace('bullet', 'decimal')})
        self.assertIn('numbered_list', self.codes(report))
        self.assertNotIn('bullet_list', self.codes(report))
        self.assertEqual(report['status'], 'ok')

    def test_empty_list_and_style_cycle(self):
        report = self.scan('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:p>', {'word/numbering.xml': numbering()})
        self.assertIn('empty_list_paragraph', self.codes(report))
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Cycle"><w:basedOn w:val="Cycle"/></w:style></w:styles>'
        report = self.scan('<w:p><w:pPr><w:pStyle w:val="Cycle"/></w:pPr></w:p>', {'word/styles.xml': styles})
        self.assertIn('unresolved_style', self.codes(report))

    def test_argument_errors_are_private(self):
        out = io.StringIO()
        with contextlib.redirect_stderr(out), self.assertRaises(SystemExit) as raised:
            main(['SYNTHETIC_PRIVATE_MARKER', '--unknown-private-marker'])
        self.assertEqual(raised.exception.code, 2)
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', out.getvalue())
        self.assertNotIn('--unknown-private-marker', out.getvalue())

    def test_metadata_custom_properties_and_non_xml(self):
        report = self.scan(extra={'docProps/custom.xml': '<Properties><property><value>SYNTHETIC_PRIVATE_MARKER</value></property></Properties>', 'word/embeddings/item.bin': b'SYNTHETIC_PRIVATE_MARKER'})
        self.assertTrue({'metadata', 'uninspected_binary'} <= self.codes(report))
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', json.dumps(report))

    def test_overrides(self):
        override = numbering().replace('<w:abstractNumId w:val="1"/>', '<w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:lvlOverride>')
        report = self.scan('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>', {'word/numbering.xml': override})
        self.assertIn('numbered_list', self.codes(report))

    def test_no_value_leak_from_malformed_xml_or_member_names(self):
        report = self.scan(extra={'SYNTHETIC_PRIVATE_MARKER.xml': '<SYNTHETIC_PRIVATE_MARKER'})
        self.assertEqual(report['status'], 'error')
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', json.dumps(report))

    def test_numbering_complexity_limit(self):
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="One"/></w:styles>'
        package(self.path, '<w:p><w:pPr><w:pStyle w:val="One"/></w:pPr></w:p>', {'word/styles.xml': styles})
        with patch('docx_preflight.scanner.MAX_NUMBERING_STEPS', 0):
            self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_unknown_xml_encoding_private_api(self):
        package(self.path, extra={'customXml/item1.xml': '<?xml version="1.0" encoding="SYNTHETIC_PRIVATE_MARKER"?><x/>'})
        out, err = io.StringIO(), io.StringIO()
        before = hashlib.sha256(self.path.read_bytes()).digest()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            report = inspect_docx(self.path)
        self.assertEqual(report['status'], 'error')
        self.assertEqual(report['findings'], [{'severity': 'error', 'code': 'unreadable_or_invalid_input', 'count': 1}])
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())
        combined = json.dumps(report) + out.getvalue() + err.getvalue()
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', combined)
        self.assertNotIn('Traceback', combined)
        self.assertEqual(out.getvalue() + err.getvalue(), '')

    def test_unknown_xml_encoding_private_cli(self):
        package(self.path, extra={'customXml/item1.xml': '<?xml version="1.0" encoding="SYNTHETIC_PRIVATE_MARKER"?><x/>'})
        for options in [[], ['--json']]:
            result = subprocess.run([sys.executable, '-B', '-m', 'docx_preflight', str(self.path), *options], capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            self.assertEqual(result.stderr, '')
            self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', result.stdout + result.stderr)
            self.assertNotIn(str(self.path), result.stdout + result.stderr)
            self.assertNotIn('Traceback', result.stdout + result.stderr)
            if options:
                self.assertEqual(json.loads(result.stdout)['status'], 'error')
            else:
                self.assertIn('unreadable_or_invalid_input', result.stdout)

    def test_numbering_lookup_work_is_linear(self):
        scanned = 0

        class CountingElement(ET.Element):
            def find(self, path, namespaces=None):
                nonlocal scanned
                scanned += len(self)
                return super().find(path, namespaces)

            def findall(self, path, namespaces=None):
                nonlocal scanned
                scanned += len(self)
                return super().findall(path, namespaces)

        def parse(text):
            return ET.fromstring(text, ET.XMLParser(target=ET.TreeBuilder(element_factory=CountingElement)))

        definitions = numbering().replace('<w:abstractNumId w:val="1"/>', '<w:unused/>' * 1000 + '<w:abstractNumId w:val="1"/>')
        paragraph = '<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        roots = {'word/numbering.xml': parse(definitions), 'word/document.xml': parse(f'<w:document xmlns:w="{W}"><w:body>{paragraph * 1000}</w:body></w:document>')}
        counts = {}

        def add(code, severity='warning'):
            counts[code] = counts.get(code, 0) + 1

        steps = _numbering(roots, add)
        self.assertEqual(counts, {'list_paragraph': 1000, 'bullet_list': 1000})
        node_count = sum(sum(1 for _ in root.iter()) for root in roots.values())
        self.assertLessEqual(scanned, node_count * 4)
        self.assertEqual(steps, node_count + 1000)

    def test_central_directory_checked_before_zipfile(self):
        with zipfile.ZipFile(self.path, 'w', zipfile.ZIP_STORED) as z:
            for i in range(257):
                z.writestr(f'item{i}', b'')
        data = bytearray(self.path.read_bytes())
        eocd = data.rindex(b'PK\x05\x06')
        # Understate both declared counts: the actual directory must be counted.
        struct.pack_into('<HH', data, eocd + 8, 1, 1)
        self.path.write_bytes(data)
        with patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('ZipInfo allocation must not start')):
            report = inspect_docx(self.path)
        self.assertEqual(report['status'], 'error')
        self.assertIn('entry_count_limit', self.codes(report))

    def test_large_forged_directory_is_rejected_before_allocation(self):
        directory_flood(self.path, 650000)
        self.assertLess(self.path.stat().st_size, scanner.MAX_ARCHIVE)
        with patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('No ZipInfo allocation')):
            report = inspect_docx(self.path)
        self.assertIn('central_directory_size_limit', self.codes(report))

    def test_unexpected_memoryerror_is_private(self):
        with patch('docx_preflight.scanner._read', side_effect=MemoryError('SYNTHETIC_PRIVATE_MARKER')):
            report = inspect_docx(self.path)
        self.assertEqual(report['status'], 'error')
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', json.dumps(report))

    def test_namespace_amplification_rejected_before_elementtree(self):
        for uri_length, count in [(65536, 1000), (262144, 1500)]:
            payload = ('<r xmlns:p="urn:synthetic:' + 'x' * uri_length + '">' + ''.join(f'<p:n{i}/>' for i in range(count)) + '</r>').encode()
            with patch('docx_preflight.scanner.ET.fromstring', side_effect=AssertionError('ElementTree must not be built')):
                with self.assertRaises(scanner.Rejected) as rejected:
                    scanner._xml(payload)
            self.assertEqual(str(rejected.exception), 'xml_namespace_uri_limit')

    def test_namespace_uri_utf8_and_utf16_limits(self):
        accepted = '<r xmlns:p="' + 'x' * scanner.MAX_NAMESPACE_URI_BYTES + '"><p:a/></r>'
        self.assertEqual(scanner._xml(accepted.encode())[0].tag, '{' + 'x' * scanner.MAX_NAMESPACE_URI_BYTES + '}a')
        for payload in ['<r xmlns="' + 'x' * (scanner.MAX_NAMESPACE_URI_BYTES + 1) + '"/>', '<r xmlns="' + '\u5408' * 1000 + '"/>', '<r xmlns="' + '&#x78;' * (scanner.MAX_NAMESPACE_URI_BYTES + 1) + '"/>']:
            for encoding in ['utf-8', 'utf-16']:
                with self.assertRaises(scanner.Rejected) as rejected:
                    scanner._xml(payload.encode(encoding))
                self.assertEqual(str(rejected.exception), 'xml_namespace_uri_limit')

    def test_xml_qualified_name_byte_limits(self):
        self.assertEqual(scanner._xml(('<' + 'a' * 512 + '/>').encode()).tag, 'a' * 512)
        for payload in ['<' + 'a' * 513 + '/>', '<r ' + 'a' * 513 + '="synthetic"/>', '<' + '\u5408' * 171 + '/>', '<r xmlns:' + 'p' * 513 + '="urn:synthetic"/>']:
            with patch('docx_preflight.scanner.ET.fromstring', side_effect=AssertionError('ElementTree must not be built')):
                with self.assertRaises(scanner.Rejected) as rejected:
                    scanner._xml(payload.encode())
            self.assertEqual(str(rejected.exception), 'xml_name_limit')

    def test_expanded_name_budget_before_elementtree(self):
        payload = b'<r xmlns:p="urn:synthetic"><p:a/><p:b/><p:c/></r>'
        with patch('docx_preflight.scanner.MAX_EXPANDED_XML_NAME_BYTES', 500), patch('docx_preflight.scanner.ET.fromstring', side_effect=AssertionError('ElementTree must not be built')):
            with self.assertRaises(scanner.Rejected) as rejected:
                scanner._xml(payload)
        self.assertEqual(str(rejected.exception), 'xml_expanded_name_limit')

    def test_actual_expanded_name_budget_with_bounded_uri(self):
        payload = ('<r xmlns:p="' + 'x' * 2048 + '">' + ''.join(f'<p:n{i}/>' for i in range(10000)) + '</r>').encode()
        with patch('docx_preflight.scanner.ET.fromstring', side_effect=AssertionError('ElementTree must not be built')):
            with self.assertRaises(scanner.Rejected) as rejected:
                scanner._xml(payload)
        self.assertEqual(str(rejected.exception), 'xml_expanded_name_limit')

    def test_expanded_name_budget_shared_across_parts(self):
        package(self.path)
        costs = []
        with zipfile.ZipFile(self.path) as z:
            for name in z.namelist():
                remaining = [scanner.MAX_EXPANDED_XML_NAME_BYTES]
                scanner._xml(z.read(name), remaining_names=remaining)
                costs.append(scanner.MAX_EXPANDED_XML_NAME_BYTES - remaining[0])
        with patch('docx_preflight.scanner.MAX_EXPANDED_XML_NAME_BYTES', max(costs)):
            self.assertIn('xml_expanded_name_limit', self.codes(inspect_docx(self.path)))

    def test_namespace_binding_scope_and_unqualified_attributes(self):
        root = scanner._xml(b'<r xmlns:p="urn:outer"><p:a/><s xmlns:p="urn:inner"><p:a/></s><p:a xml:lang="en"/></r>')
        self.assertEqual(root[0].tag, '{urn:outer}a')
        self.assertEqual(root[1][0].tag, '{urn:inner}a')
        self.assertEqual(root[2].tag, '{urn:outer}a')
        self.assertIn('{http://www.w3.org/XML/1998/namespace}lang', root[2].attrib)
        root = scanner._xml(b'<r xmlns="urn:outer"><a attr="synthetic"><b xmlns="" attr="synthetic"/></a></r>')
        self.assertEqual(root[0].tag, '{urn:outer}a')
        self.assertEqual(root[0][0].tag, 'b')
        self.assertEqual(root[0].attrib, {'attr': 'synthetic'})

    def test_namespace_rejection_api_cli_privacy(self):
        payload = '<r xmlns:p="urn:SYNTHETIC_PRIVATE_MARKER:' + 'x' * 65536 + '"><p:a/></r>'
        package(self.path, extra={'customXml/item.xml': payload}, compression=zipfile.ZIP_STORED)
        before = hashlib.sha256(self.path.read_bytes()).digest()
        report = inspect_docx(self.path)
        self.assertIn('xml_namespace_uri_limit', self.codes(report))
        self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', json.dumps(report))
        for options in [[], ['--json']]:
            result = subprocess.run([sys.executable, '-B', '-m', 'docx_preflight', str(self.path), *options], capture_output=True, text=True)
            self.assertEqual(result.returncode, 2)
            self.assertEqual(result.stderr, '')
            self.assertNotIn('SYNTHETIC_PRIVATE_MARKER', result.stdout)
            self.assertNotIn(str(self.path), result.stdout)
            if options:
                self.assertEqual(json.loads(result.stdout)['status'], 'error')
        self.assertEqual(before, hashlib.sha256(self.path.read_bytes()).digest())

    def test_numbering_indexes_style_and_abstract_children(self):
        paragraph = '<w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        styles = f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Base">' + '<w:unused/>' * 1000 + '<w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style></w:styles>'
        definitions = numbering().replace('<w:lvl w:ilvl="0">', '<w:unused/>' * 1000 + '<w:lvl w:ilvl="0">')
        roots = {'word/styles.xml': ET.fromstring(styles), 'word/numbering.xml': ET.fromstring(definitions), 'word/document.xml': ET.fromstring(f'<w:document xmlns:w="{W}"><w:body>{paragraph * 1000}</w:body></w:document>')}
        counts = {}
        steps = _numbering(roots, lambda code, severity='warning': counts.update({code: counts.get(code, 0) + 1}))
        self.assertEqual(counts, {'list_paragraph': 1000, 'bullet_list': 1000})
        self.assertEqual(steps, sum(sum(1 for _ in r.iter()) for r in roots.values()) + 2000)

    def test_numbering_budget_counts_ignored_elements(self):
        package(self.path, extra={'word/numbering.xml': numbering().replace('<w:abstractNumId w:val="1"/>', '<w:unused/>' * 100 + '<w:abstractNumId w:val="1"/>')})
        with patch('docx_preflight.scanner.MAX_NUMBERING_STEPS', 50):
            self.assertIn('numbering_complexity_limit', self.codes(inspect_docx(self.path)))

    def test_shared_long_ids_are_not_revalidated_per_paragraph(self):
        work = 0

        class CountingString(str):
            __hash__ = str.__hash__

            def isdecimal(self):
                nonlocal work
                work += len(self)
                return super().isdecimal()

            def __eq__(self, other):
                nonlocal work
                if isinstance(other, str) and len(self) == len(other):
                    work += len(self)
                return super().__eq__(other)

        long_style = 'SyntheticStyle' * 100
        long_num = '7' * 1000
        styles = ET.fromstring(f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style></w:styles>')
        styles[0].set('{' + W + '}styleId', CountingString(long_style))
        styles[0][0][0][0].set('{' + W + '}val', CountingString(long_num))
        styles[1][0].set('{' + W + '}val', CountingString(long_style))
        definitions = ET.fromstring(numbering())
        definitions[1].set('{' + W + '}numId', CountingString(long_num))
        paragraph = '<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        document = ET.fromstring(f'<w:document xmlns:w="{W}"><w:body>{paragraph * 1000}</w:body></w:document>')
        counts = {}
        _numbering({'word/styles.xml': styles, 'word/numbering.xml': definitions, 'word/document.xml': document}, lambda code, severity='warning': counts.update({code: counts.get(code, 0) + 1}))
        self.assertEqual(counts, {'list_paragraph': 1000, 'bullet_list': 1000})
        self.assertLess(work, 20 * (len(long_style) + len(long_num)))

    def test_duplicate_levels_and_override_fallback(self):
        paragraph = '<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        definitions = numbering().replace('</w:abstractNum>', '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>')
        definitions = definitions.replace('<w:abstractNumId w:val="1"/>', '<w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:lvlOverride><w:lvlOverride w:ilvl="0"><w:startOverride w:val="3"/></w:lvlOverride>')
        self.assertIn('numbered_list', self.codes(self.scan(paragraph, {'word/numbering.xml': definitions})))
        missing = definitions.replace('<w:startOverride w:val="3"/>', '<w:lvl w:ilvl="0"/>')
        self.assertIn('unresolved_numbering', self.codes(self.scan(paragraph, {'word/numbering.xml': missing})))

    def test_total_xml_and_attribute_limits(self):
        package(self.path)
        with patch('docx_preflight.scanner.MAX_TOTAL_XML_NODES', 3):
            self.assertIn('total_xml_complexity_limit', self.codes(inspect_docx(self.path)))
        attributes = ' '.join(f'a{i}="synthetic"' for i in range(257))
        self.assertIn('xml_attribute_limit', self.codes(self.scan(extra={'customXml/item1.xml': f'<x {attributes}/>'})))
        package(self.path, extra={'customXml/item1.xml': '<x a="synthetic" b="synthetic"/>'})
        with patch('docx_preflight.scanner.MAX_TOTAL_XML_ATTRIBUTES', 1):
            self.assertIn('total_xml_complexity_limit', self.codes(inspect_docx(self.path)))
        with patch('docx_preflight.scanner.MAX_TOTAL_XML_BYTES', 1):
            self.assertIn('total_xml_size_limit', self.codes(inspect_docx(self.path)))

    def test_directory_rechecked_after_input_changes(self):
        package(self.path)
        original_check = scanner._check_directory
        calls = 0

        def mutate_after_first_check(source, size):
            nonlocal calls
            original_check(source, size)
            calls += 1
            if calls == 1:
                data = bytearray(self.path.read_bytes())
                eocd = data.rindex(b'PK\x05\x06')
                struct.pack_into('<I', data, eocd + 12, 0xffff0000)
                self.path.write_bytes(data)
                # Discard any old BufferedReader bytes so the snapshot observes
                # this synthetic same-size change, rather than a valid old copy.
                source.seek(size + 1)
                source.seek(0)

        with patch('docx_preflight.scanner._check_directory', side_effect=mutate_after_first_check), patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('No ZipInfo allocation')):
            self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_zip_reader_uses_validated_snapshot(self):
        package(self.path)
        original_zip = zipfile.ZipFile

        def mutate_after_snapshot(source):
            self.path.write_bytes(b'synthetic external change')
            return original_zip(source)

        with patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=mutate_after_snapshot):
            self.assertEqual(inspect_docx(self.path)['status'], 'ok')

    def test_directory_size_limit_before_allocation(self):
        package(self.path)
        with patch('docx_preflight.scanner.MAX_DIRECTORY', 1), patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('No ZipInfo allocation')):
            self.assertIn('central_directory_size_limit', self.codes(inspect_docx(self.path)))

    def test_directory_count_mismatch(self):
        package(self.path)
        data = bytearray(self.path.read_bytes())
        eocd = data.rindex(b'PK\x05\x06')
        struct.pack_into('<HH', data, eocd + 8, 1, 1)
        self.path.write_bytes(data)
        self.assertIn('inconsistent_directory_count', self.codes(inspect_docx(self.path)))

    def test_directory_hostile_offsets_sizes_and_lengths(self):
        for location, displacement, format_code, value in [('end', 16, '<I', 1), ('end', 12, '<I', 1), ('central', 42, '<I', 0xffffff00), ('central', 28, '<H', 65534), ('central', 30, '<H', 65534), ('central', 32, '<H', 65534)]:
            package(self.path)
            data = bytearray(self.path.read_bytes())
            location = data.rindex(b'PK\x05\x06') if location == 'end' else data.index(b'PK\x01\x02')
            struct.pack_into(format_code, data, location + displacement, value)
            self.path.write_bytes(data)
            with patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('No ZipInfo allocation')):
                self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_zip64_and_multidisk_are_explicitly_rejected(self):
        for displacement, fmt, value, code in [(8, '<H', 65535, 'unsupported_zip64'), (12, '<I', 0xffffffff, 'unsupported_zip64'), (16, '<I', 0xffffffff, 'unsupported_zip64'), (4, '<H', 1, 'unsupported_multidisk_or_count')]:
            package(self.path)
            data = bytearray(self.path.read_bytes())
            eocd = data.rindex(b'PK\x05\x06')
            struct.pack_into(fmt, data, eocd + displacement, value)
            self.path.write_bytes(data)
            self.assertIn(code, self.codes(inspect_docx(self.path)))
        with zipfile.ZipFile(self.path, 'w') as z:
            with z.open('synthetic', 'w', force_zip64=True) as stream:
                stream.write(b'synthetic')
        self.assertIn('unsupported_zip64_or_version', self.codes(inspect_docx(self.path)))

    def test_zip64_locator_and_extra_fields(self):
        package(self.path)
        data = self.path.read_bytes()
        eocd = data.rindex(b'PK\x05\x06')
        self.path.write_bytes(data[:eocd] + b'PK\x06\x07' + bytes(16) + data[eocd:])
        self.assertIn('unsupported_zip64', self.codes(inspect_docx(self.path)))
        with zipfile.ZipFile(self.path, 'w') as z:
            info = zipfile.ZipInfo('synthetic')
            info.extra = struct.pack('<HH', 1, 0)
            z.writestr(info, b'')
        self.assertIn('unsupported_zip64', self.codes(inspect_docx(self.path)))

    def test_zip_comments_unicode_and_benign_extra(self):
        package(self.path)
        with zipfile.ZipFile(self.path, 'a') as z:
            z.comment = b'synthetic comment'
            info = zipfile.ZipInfo('word/media/\u5408\u6210.bin')
            info.extra = struct.pack('<HH', 0xffff, 4) + b'test'
            z.writestr(info, b'synthetic')
        self.assertIn('uninspected_binary', self.codes(inspect_docx(self.path)))

    def test_truncated_directory_and_local_header(self):
        package(self.path)
        original = self.path.read_bytes()
        for data in [original[:-1], b'BAD!' + original[4:]]:
            self.path.write_bytes(data)
            with patch('docx_preflight.scanner.zipfile.ZipFile', side_effect=AssertionError('No ZipInfo allocation')):
                self.assertEqual(inspect_docx(self.path)['status'], 'error')

    def test_cli_codes_json_and_no_filename(self):
        package(self.path)
        for args, expected in [([str(self.path), '--json'], 0), ([str(self.path), '--json', '--fail-on', 'info'], 0)]:
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                self.assertEqual(main(args), expected)
            self.assertEqual(json.loads(out.getvalue())['status'], 'ok')
            self.assertNotIn(str(self.path), out.getvalue())
        self.scan('<w:p><w:ins/></w:p>')
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(main([str(self.path)]), 1)
            self.assertEqual(main([str(self.path), '--fail-on', 'none']), 0)
            self.assertEqual(main([str(self.path.parent / 'missing.docx'), '--json']), 2)


if __name__ == '__main__':
    unittest.main()
