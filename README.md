# docx-preflight

**Review Word list settings and sharing risks before you send a document.**

[English](README.md) · [日本語](README.ja.md)

Unexpected bullet marks? An author name or an old comment still in a draft? **docx-preflight** gives writers, editors, and developers a local checklist of OOXML structures worth reviewing. Normal bullets and numbered lists are reported as information; suspicious or unresolved settings and privacy hints are warnings.

**Python 3.10+ · No runtime dependencies · Reads without changing the document · MIT**

Use it for a quick review of a report or draft, or consume its JSON and exit codes in a local workflow. It inspects structure, not the meaning of your text. **A completed check is not a guarantee that the document contains no private information.**

[Quick start](#quick-start) · [Read the output](#read-output) · [Checks and limits](#checks) · [Privacy](#privacy) · [Diagnostic codes](#diagnostics) · [Exit codes](#exit-codes) · [Contributing](#contributing)

<a id="quick-start"></a>

## Quick start

Download this repository and open a terminal in its folder. Check that your interpreter is Python 3.10 or later, then replace `document.docx` with your file's path:

```sh
python --version
python -m docx_preflight document.docx
```

That's enough to inspect a file: source-directory use needs no installation. Put paths with spaces in quotes. If your system uses `python3` or `py -3`, use that interpreter in place of `python` throughout these examples.

To install the command from the downloaded source:

```sh
python -m pip install .
docx-preflight "document.docx"
```

Installation may obtain the setuptools build dependency. Inspection itself has no network code. This project is not published to PyPI; use the source installation above.

For scripts or a stricter review threshold:

```sh
python -m docx_preflight document.docx --json
python -m docx_preflight document.docx --fail-on info
python -m docx_preflight --help
```

<a id="read-output"></a>

## Read the output

For a synthetic file with two valid bullet paragraphs and one populated author property:

```text
DOCX preflight: warnings
Heuristic OOXML checks; not a privacy or rendering guarantee.
info: bullet_list (2)
info: list_paragraph (2)
warning: metadata (1)
```

- The two `info` lines describe the same two paragraphs. They are not four separate problems, and valid bullets do not automatically mean an error.
- `metadata (1)` means one populated property needs review. The author's value is withheld; a populated property can be harmless.
- `warnings` means review the flagged structures. This example returns exit code **1** with the default threshold.

The `count` for each code is the number of matching structures. Different codes can overlap; do not sum them as a number of people or affected paragraphs. Results are aggregated across the package, without locations or document values.

<details>
<summary>Try the example with a synthetic file</summary>

Save this Python snippet as `make_demo.py` in the source folder and run `python make_demo.py`. It creates a minimal synthetic DOCX for inspecting XML, not a sample whose layout has been verified in Word. It refuses to overwrite an existing `synthetic-demo.docx`.

```python
from zipfile import ZIP_DEFLATED, ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/package/2006/content-types"
parts = {
    "[Content_Types].xml": f'<Types xmlns="{C}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    "_rels/.rels": f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="{O}/officeDocument" Target="word/document.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    "word/document.xml": f'<w:document xmlns:w="{W}"><w:body>' + ''.join('<w:p><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr><w:r><w:t>' + text + '</w:t></w:r></w:p>' for text in ["First synthetic item", "Second synthetic item"]) + '</w:body></w:document>',
    "word/numbering.xml": f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#x2022;"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="1"/></w:num></w:numbering>',
    "word/_rels/document.xml.rels": f'<Relationships xmlns="{R}"><Relationship Id="r1" Type="{O}/numbering" Target="numbering.xml"/></Relationships>',
    "docProps/core.xml": '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Synthetic Author</dc:creator></cp:coreProperties>',
}
try:
    with ZipFile("synthetic-demo.docx", "x", ZIP_DEFLATED) as archive:
        for name, text in parts.items():
            archive.writestr(name, text)
except FileExistsError:
    raise SystemExit("The synthetic demo already exists. Choose a new demo filename.")
```

```sh
python -m docx_preflight synthetic-demo.docx
python -m docx_preflight synthetic-demo.docx --json
```

</details>

With `--json`, the same example produces the following data (formatted here for readability):

```json
{
  "schema_version": 1,
  "status": "warnings",
  "findings": [
    {"severity": "info", "code": "bullet_list", "count": 2},
    {"severity": "info", "code": "list_paragraph", "count": 2},
    {"severity": "warning", "code": "metadata", "count": 1}
  ],
  "scope": "Heuristic OOXML checks; not a privacy or rendering guarantee."
}
```

<a id="checks"></a>

## What it checks, and what needs your judgment

| Your question | What the tool provides | What you still review |
| --- | --- | --- |
| Why might a paragraph show an unwanted bullet? | Effective numbering, style inheritance, empty list paragraphs, unresolved definitions, and numbering inherited from defaults. | Whether a marker is intended and how Word renders it. |
| Could a draft reveal collaboration information? | Hints for populated properties, comments, selected revisions, identities, external relationships, fields, and hidden-text settings. | The actual document and whether its contents are appropriate to share. |
| Can I integrate this into a local check? | JSON, fixed codes/counts, and configurable failure thresholds. | Your workflow's interpretation of those findings. |

The scanner merges document defaults, paragraph style inheritance, and direct properties. It resolves numbering levels and inline level overrides, respects `numId=0`, and ignores unused numbering definitions.

It does not repair bullet settings, remove metadata, accept/reject revisions, render Word pages, or find every personal fact. It cannot inspect image metadata, OCR, embedded binaries, visible prose, visual concealment, or every OOXML extension. Advanced style-linked numbering and layout-specific behavior need manual review. Both `ok` and `warnings` are completed inspections; `error` means the result is incomplete.

<a id="privacy"></a>

## Your document stays local

Inspection never uploads documents, fetches external links, opens Office, extracts ZIP members to disk, or changes inputs. There is no cleanup command. Only fixed codes, severity and counts leave the reader: filenames/paths, member names, author values, comment contents, link targets, document text, and raw XML are withheld, including for rejected input and command argument errors. Your shell or external logging system may still record command arguments.

Review the document and, where useful, run [Microsoft's Document Inspector](https://support.microsoft.com/en-us/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks) on a separate copy. It also has detection limits. Never attach private documents to an issue report.

<a id="diagnostics"></a>

## Diagnostic reference

### List settings

| Code | Severity | Meaning |
| --- | --- | --- |
| `list_paragraph` | info | A paragraph has a resolved, visible numbering format; review whether the list is intended. |
| `bullet_list` / `numbered_list` | info | Bullet format / another visible numbering format, respectively. |
| `suppressed_numbering` | info | Numbering format is `none`. |
| `empty_list_paragraph` | warning | Numbered paragraph has no non-whitespace text; it may produce an unwanted marker. Drawing-only items can also trigger this. |
| `default_numbering` | warning | Numbering reaches a paragraph through document or default paragraph-style settings; review scope. |
| `unresolved_numbering` | warning | Invalid ID/level or missing numbering/level format. |
| `unresolved_style` | warning | A referenced style is missing or the inheritance chain cycles. |
| `unsupported_numbering_link` | warning | Numbering depends on a numbering-style link that this tool does not resolve. |

These are heuristics: XML cannot tell whether an author intended a list. Valid bullets and numbered lists are information, not automatic errors. This is not a Word rendering engine. Advanced style-linked numbering, implicit level selection from style associations, compatibility modes, picture bullets, and layout-specific behavior need manual review.

### Privacy hints

| Code | Meaning |
| --- | --- |
| `metadata` | Populated entries in `docProps` (including creator, last editor, dates, application properties, and custom properties). Counts may include benign properties. |
| `comments` / `comment_reference` | Comment parts / references, including common extended comment parts. |
| `tracked_changes` | Insertions, deletions, moves, property changes, and selected table/numbering revisions. |
| `collaborator_data` / `collaborator_identity` | Common people parts / author, initials, user ID, or email attributes. |
| `external_relationship` | A relationship has external target mode; no target is followed. |
| `field_code` | Field instructions; these may contain local paths or external references. All fields are flagged conservatively. |
| `hidden_text` | Explicit `vanish` or `webHidden` settings that are enabled. |
| `custom_xml` | Custom XML parts. |
| `uninspected_binary` | Binary/non-XML parts such as images and embedded objects; their contents and image metadata are not inspected. |

Every privacy hint is a warning. A warning means review is needed; it does not establish that private information exists.

**No findings is not a safety certificate.** Private material can remain in visible prose, headers/footers, arbitrary extension XML, images, or embedded content even when `status` is `ok`.

<a id="exit-codes"></a>

## Exit codes

| Exit | Meaning |
| --- | --- |
| 0 | No finding at the selected threshold (default: warning). Never a complete privacy guarantee. |
| 1 | At least one finding meets `--fail-on warning` or `--fail-on info`. |
| 2 | Inspection failed, input was rejected, or command arguments were invalid. Results are incomplete; do not interpret this as clean. |

`--fail-on none` returns 0 for completed inspections even with findings; inspection errors always return 2. JSON `status` is `ok`, `warnings`, or `error`, independent of the threshold.

<details>
<summary>Input limits and parser scope</summary>

Oversized or unsupported input produces `status: error` and exit code 2. Use these limits to understand a rejection, rather than treating a failed inspection as a clean document.

The reader limits the archive to 32 MiB, 256 entries, 8 MiB per expanded part, 64 MiB total expanded size, and 200:1 compression ratio per part. XML is limited to 16 MiB total, 100 levels, 100,000 elements per part, 200,000 elements across all parts, 256 attributes per element, and 200,000 attributes across all parts. Oversized files are rejected, not partially declared clean. Only stored/deflated, unencrypted ZIP entries are supported. Duplicate names (case insensitive), unsafe names, symlink entries, CRC errors, malformed XML, DTDs, and entity declarations are rejected. XML declarations specifying encoding are allowed; unknown encoding errors use a fixed diagnostic. Both UTF-8 and UTF-16 DTD rejection are tested. Parts are read in memory and never extracted.

Before `ZipFile` creates member objects, a bounded ZIP32 header reader caps the central directory at 256 KiB and walks its actual records, including when EOCD counts are forged. It validates declared sizes, variable-field lengths, local header offsets/names, and directory placement. ZIP64, multi-disk archives, extraction versions 4.5+, noncanonical offsets, trailing bytes, and ambiguous end records are rejected. This intentionally supports a narrower set of archives than Python's general ZIP reader. A bounded immutable snapshot is checked again before use, so a file changing during inspection cannot replace the validated directory with a larger one.

Before ElementTree builds any tree, an Expat pass without namespace expansion limits decoded namespace URIs to 2,048 UTF-8 bytes and qualified element/attribute names (including namespace declarations) to 512 UTF-8 bytes. It tracks scoped namespace bindings without copying whole scope maps and budgets the names that ElementTree would expand. All parts share a conservative 64 MiB name budget: `4 * (characters + 2) + 80` per expanded element/attribute name or namespace declaration (URI included). Repeated occurrences count too. This prevents a short input with one long URI and many distinct names from creating disproportionately large tree names. Namespace rebindings, default namespace removal, `xml:` attributes, UTF-16, and character references are covered by synthetic tests. URI/name/expansion limit errors use fixed codes without showing values.

Numbering resolution is limited to 200,000 steps across initial XML indexing, paragraphs, and style-chain edges per inspection. Ignored XML children also count. Shared number definitions, levels, overrides, first-child properties, and text flags are indexed once; long reference strings and numeric classifications are cached once. Resolution takes work proportional to the indexed XML plus the traversed style edges, instead of paragraphs times shared-definition size. Unusually complex documents are rejected. These bounds are not a hard memory or wall-clock guarantee under every environment.

Only packages with a standard DOCX content type and the main part at `word/document.xml` are accepted. Macro-enabled main types and other main-part locations are rejected. Transitional and Strict WordprocessingML namespace elements are handled, but full OOXML schema validation is outside scope. This is a bounded parser, not a malware scanner or a sandbox.

</details>

<a id="contributing"></a>

## Development and contributing

Python **3.10+** is required. Local verification has run on Python 3.13; the CI configuration targets Python 3.10 and 3.13 on Windows and Linux. No runtime packages are required to run the source or its tests.

```sh
python -m unittest discover -s tests -v
python -m docx_preflight --help
```

Tests create synthetic ZIP/XML inputs in temporary directories, verify input hashes remain unchanged, and exercise list inheritance, privacy output suppression, malformed/encrypted/oversized ZIPs, CRC, duplicates, XML declarations, linear numbering lookup work, forged directory counts, ZIP64 rejection, namespace/name limits, and input changes. No real documents or personal samples are included. CI is configured for Python 3.10 and 3.13 on Windows and Linux. ZIP32 header interpretation follows the [PKWARE format specification](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT); the specification itself is not bundled.

For a bug report, include the diagnostic codes, expected behavior, Python version, and a small synthetic reproduction. Redact private paths from external logs. For a change, explain its behavior, add a meaningful synthetic regression test when needed, and keep both READMEs consistent. Documentation fixes and additional translations are welcome; add a language link when introducing a translation. Do not include personal documents, real author values, secrets, or credentials in examples.

Related projects: [python-docx](https://github.com/python-openxml/python-docx) creates/edits Word documents; [mat2](https://github.com/jvoisin/mat2) removes metadata across formats. This project's narrow purpose is read-only diagnostic output for lists and sharing review. No source or samples were copied from those projects.

## License

[MIT](LICENSE). See the license file for the terms.
