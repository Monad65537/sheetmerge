"""Validate worksheet structure before decoding any values with openpyxl."""

import io
import posixpath
import re
import zlib
from dataclasses import dataclass
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from .config import InputLimits

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MAX_ROW = 1_048_576
MAX_COLUMN = 16_384


class XlsxStructureError(ValueError):
    """Ambiguous or unsupported worksheet structure must not be ingested."""


@dataclass(frozen=True)
class StoredCell:
    data_type: str
    value_text: str | None
    has_formula: bool


class _XmlBudget(ET.TreeBuilder):
    """Reject declarations and excessive nesting/element counts while parsing."""

    def __init__(self, limits: InputLimits, part: str):
        super().__init__()
        self.limits = limits
        self.part = part
        self.depth = 0
        self.elements = 0

    def start(self, tag, attributes):
        self.depth += 1
        self.elements += 1
        self.limits.check("max_xml_depth", self.depth, self.part)
        self.limits.check("max_xml_elements", self.elements, self.part)
        return super().start(tag, attributes)

    def end(self, tag):
        element = super().end(tag)
        self.depth -= 1
        return element

    def doctype(self, name, public_id, system_id):
        raise XlsxStructureError(f"{self.part}: XML DTD/entity declarations are not allowed")


def _parse_xml(data: bytes, part: str, limits: InputLimits, *, required: bool = True):
    try:
        return ET.fromstring(data, parser=ET.XMLParser(target=_XmlBudget(limits, part)))
    except ET.ParseError as exc:
        if required:
            raise XlsxStructureError(f"{part}: malformed XML: {exc}") from exc
        return None  # Non-XML binary parts need not be well-formed XML.


def _checked_members(archive: ZipFile, limits: InputLimits) -> dict[str, bytes]:
    """Check advertised budgets first, then bounded inflation, CRCs, and XML."""
    entries = archive.infolist()
    limits.check("max_zip_members", len(entries), "XLSX archive")
    if len(entries) != len({entry.filename for entry in entries}):
        raise XlsxStructureError("Duplicate ZIP member names are not allowed")
    declared_total = 0
    for entry in entries:
        name = entry.filename
        original_name = getattr(entry, "orig_filename", name)
        canonical = posixpath.normpath(name.rstrip("/"))
        if (
            original_name != name
            or "\x00" in original_name
            or not name
            or name.startswith("/")
            or "\\" in name
            or ":" in name
            or (canonical != name.rstrip("/") or canonical == ".." or canonical.startswith("../"))
        ):
            raise XlsxStructureError(f"Unsafe ZIP member path: {original_name!r}")
        if entry.flag_bits & 1:
            raise XlsxStructureError(f"{name}: encrypted ZIP members are not supported")
        if entry.compress_type not in (ZIP_STORED, ZIP_DEFLATED):
            raise XlsxStructureError(f"{name}: unsupported ZIP compression method")
        limits.check("max_zip_member_bytes", entry.file_size, name)
        declared_total += entry.file_size
        limits.check("max_zip_uncompressed_bytes", declared_total, "XLSX archive")
        if entry.file_size:
            if not entry.compress_size:
                raise XlsxStructureError(f"{name}: nonempty ZIP member has zero compressed size")
            ratio = (entry.file_size + entry.compress_size - 1) // entry.compress_size
            limits.check("max_zip_compression_ratio", ratio, name)

    contents = {}
    actual_total = 0
    for entry in entries:
        chunks = []
        size = 0
        try:
            with archive.open(entry) as stream:
                while True:
                    chunk = stream.read(
                        min(
                            65_536,
                            limits.max_zip_member_bytes - size + 1,
                            limits.max_zip_uncompressed_bytes - actual_total + 1,
                        )
                    )
                    if not chunk:
                        break
                    size += len(chunk)
                    actual_total += len(chunk)
                    limits.check("max_zip_member_bytes", size, entry.filename)
                    limits.check("max_zip_uncompressed_bytes", actual_total, "XLSX archive")
                    chunks.append(chunk)
        except (BadZipFile, RuntimeError, EOFError, zlib.error) as exc:
            raise XlsxStructureError(f"{entry.filename}: malformed ZIP member: {exc}") from exc
        if size != entry.file_size:
            raise XlsxStructureError(f"{entry.filename}: ZIP member size does not match metadata")
        data = b"".join(chunks)
        if not entry.is_dir():
            # Scan nonstandard part names too: a valid XML document must receive
            # the same DTD/depth/element guards regardless of its file extension.
            _parse_xml(
                data,
                entry.filename,
                limits,
                required=entry.filename.lower().endswith((".xml", ".rels")),
            )
        contents[entry.filename] = data
    return contents


def _worksheet(
    data: bytes, name: str, limits: InputLimits, header_row: int, records_before: int
) -> tuple[bytes, dict[str, StoredCell], int]:
    root = _parse_xml(data, f"sheet {name!r}", limits)
    tables = root.findall(f"{{{MAIN}}}sheetData")
    if root.tag != f"{{{MAIN}}}worksheet" or len(tables) != 1:
        raise XlsxStructureError(f"sheet {name!r}: expected one worksheet/sheetData element")
    stored = {}
    previous_row = 0
    row_count = 0
    for row in tables[0]:
        if row.tag != f"{{{MAIN}}}row":
            raise XlsxStructureError(f"sheet {name!r}: unexpected element inside sheetData")
        row_text = row.get("r", str(previous_row + 1))
        if not re.fullmatch(r"[0-9]+", row_text) or not 1 <= int(row_text) <= MAX_ROW:
            raise XlsxStructureError(f"sheet {name!r}: invalid row coordinate {row_text!r}")
        row_number = int(row_text)
        limits.check("max_worksheet_rows", row_number, f"sheet {name!r}")
        limits.check(
            "max_total_records",
            records_before + max(0, row_number - header_row),
            f"sheet {name!r}, row {row_number}",
        )
        if row_number <= previous_row:
            raise XlsxStructureError(f"sheet {name!r}: duplicate or out-of-order row {row_number}")
        row.set("r", str(row_number))
        previous_row = row_number
        previous_column = 0
        row_count += 1
        for cell in row:
            if cell.tag == f"{{{MAIN}}}extLst":
                continue
            if cell.tag != f"{{{MAIN}}}c":
                raise XlsxStructureError(f"sheet {name!r}, row {row_number}: unexpected row child")
            reference = cell.get("r")
            if reference is None:
                column = previous_column + 1
            else:
                match = re.fullmatch(r"([A-Za-z]{1,3})([0-9]+)", reference)
                if match is None:
                    raise XlsxStructureError(
                        f"sheet {name!r}: invalid cell coordinate {reference!r}"
                    )
                column = column_index_from_string(match[1].upper())
                if int(match[2]) != row_number:
                    raise XlsxStructureError(
                        f"sheet {name!r}: cell {reference!r} does not belong to row {row_number}"
                    )
            if not 1 <= column <= MAX_COLUMN:
                raise XlsxStructureError(f"sheet {name!r}, row {row_number}: column out of bounds")
            limits.check("max_worksheet_columns", column, f"sheet {name!r}, row {row_number}")
            coordinate = f"{get_column_letter(column)}{row_number}"
            if coordinate in stored:
                raise XlsxStructureError(f"sheet {name!r}: duplicate cell coordinate {coordinate}")
            if column <= previous_column:
                raise XlsxStructureError(f"sheet {name!r}: out-of-order cell {coordinate}")
            previous_column = column
            cell.set("r", coordinate)
            values = cell.findall(f"{{{MAIN}}}v")
            formulas = cell.findall(f"{{{MAIN}}}f")
            inline = cell.findall(f"{{{MAIN}}}is")
            data_type = cell.get("t", "n")
            if (
                data_type not in ("n", "b", "e", "s", "str", "inlineStr", "d")
                or len(values) > 1
                or len(formulas) > 1
                or len(inline) > 1
                or (inline and (values or formulas or data_type != "inlineStr"))
                or (data_type == "inlineStr" and values)
                or any(list(value) for value in values)
                or any(
                    child.tag
                    not in {f"{{{MAIN}}}v", f"{{{MAIN}}}f", f"{{{MAIN}}}is", f"{{{MAIN}}}extLst"}
                    for child in cell
                )
            ):
                raise XlsxStructureError(f"sheet {name!r}: ambiguous/unsupported cell {coordinate}")
            stored[coordinate] = StoredCell(
                data_type, values[0].text if values else None, bool(formulas)
            )
            if data_type == "d":
                # Preserve ISO text for our strict validator, rather than letting
                # openpyxl truncate fractions or accept a valid prefix plus junk.
                # This changes only the in-memory decoder copy, never the source.
                cell.set("t", "str")
    if sum(1 for _ in root.iter(f"{{{MAIN}}}row")) != row_count or sum(
        1 for _ in root.iter(f"{{{MAIN}}}c")
    ) != len(stored):
        raise XlsxStructureError(f"sheet {name!r}: row/cell outside its expected parent")
    # XML parsers normalize literal CR characters; keep original string content.
    decoded_copy = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return decoded_copy.replace(b"\r", b"&#13;"), stored, max(0, previous_row - header_row)


def validate_xlsx(
    snapshot: bytes,
    sheets: tuple[str, ...],
    limits: InputLimits,
    *,
    header_row: int = 1,
    records_before: int = 0,
) -> tuple[bytes, dict[str, dict[str, StoredCell]]]:
    """Return a validated decoder copy and its unique coordinate/value index."""
    with ZipFile(io.BytesIO(snapshot)) as archive:
        contents = _checked_members(archive, limits)
        workbook = _parse_xml(contents["xl/workbook.xml"], "xl/workbook.xml", limits)
        relationships = {}
        for element in _parse_xml(
            contents["xl/_rels/workbook.xml.rels"], "xl/_rels/workbook.xml.rels", limits
        ):
            identifier = element.attrib["Id"]
            if identifier in relationships:
                raise XlsxStructureError(f"Duplicate workbook relationship {identifier!r}")
            relationships[identifier] = element
        sheet_elements = workbook.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet")
        names = [sheet.attrib["name"] for sheet in sheet_elements]
        if len({name.casefold() for name in names}) != len(names):
            raise XlsxStructureError("Duplicate worksheet names are not allowed")
        missing = set(sheets) - set(names)
        if missing:
            raise XlsxStructureError(f"Configured sheets are missing: {sorted(missing)}")
        replacements = {}
        cells = {}
        for sheet in sheet_elements:
            name = sheet.attrib["name"]
            if name not in sheets:
                continue
            relationship = relationships[sheet.attrib[f"{{{DOCUMENT_REL}}}id"]]
            if relationship.attrib.get("TargetMode") == "External" or not relationship.attrib[
                "Type"
            ].endswith("/worksheet"):
                raise XlsxStructureError(
                    f"sheet {name!r}: expected an internal worksheet relationship"
                )
            target = relationship.attrib["Target"]
            part = posixpath.normpath(
                target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
            )
            if part.startswith("../") or part in replacements:
                raise XlsxStructureError(f"sheet {name!r}: invalid or reused worksheet part")
            replacements[part], cells[name], records = _worksheet(
                contents[part], name, limits, header_row, records_before
            )
            records_before += records
        output = io.BytesIO()
        with ZipFile(output, "w") as target:
            for entry in archive.infolist():
                data = replacements.get(entry.filename)
                target.writestr(entry, contents[entry.filename] if data is None else data)
    return output.getvalue(), cells
