from __future__ import annotations

import binascii
import stat
import struct
import zipfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pytest

import formulafence.workbook as workbook_module
from formulafence.cli import main
from formulafence.models import WorkbookLoadError
from formulafence.workbook import load_snapshot

from .helpers import make_model


def _append_member(
    path: Path,
    name: str,
    payload: bytes,
    *,
    compression: int = ZIP_DEFLATED,
) -> None:
    with ZipFile(path, "a", compression=compression) as archive:
        archive.writestr(name, payload)


def _last_central_directory_offset(contents: bytes | bytearray) -> int:
    offset = contents.rfind(b"PK\x01\x02")
    assert offset >= 0
    return offset


def _reject_before_workbook_readers(monkeypatch: pytest.MonkeyPatch, path: Path) -> str:
    def unexpected_reader(*args, **kwargs):
        raise AssertionError("a workbook reader ran before archive preflight rejected input")

    monkeypatch.setattr(workbook_module, "_workbook_tab_order_metadata", unexpected_reader)
    with pytest.raises(WorkbookLoadError, match="safety preflight") as error:
        load_snapshot(path)
    return str(error.value)


def test_archive_preflight_accepts_an_ordinary_workbook(tmp_path: Path) -> None:
    workbook = make_model(tmp_path / "ordinary.xlsx")

    snapshot = load_snapshot(workbook)

    assert snapshot.file_type == "xlsx"
    assert set(snapshot.sheets) == {"Inputs", "Model", "Dashboard", "Control"}


def test_archive_preflight_accepts_a_valid_zip64_workbook(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(zipfile, "ZIP64_LIMIT", 1)
    workbook = make_model(tmp_path / "zip64.xlsx")

    assert bytes((80, 75, 6, 6)) in workbook.read_bytes()
    assert load_snapshot(workbook).file_type == "xlsx"


def test_archive_preflight_rejects_source_size_before_any_zip_reader(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_SOURCE_BYTES", 1)

    def unexpected_zip_reader(*args, **kwargs):
        raise AssertionError("ZipFile ran before the source-size safety gate")

    monkeypatch.setattr(workbook_module, "ZipFile", unexpected_zip_reader)

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_excessive_entry_count_before_zip_reader(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "too-many-members.xlsx")
    with ZipFile(workbook) as archive:
        member_count = len(archive.infolist())
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_ENTRY_COUNT", member_count - 1)

    def unexpected_zip_reader(*args, **kwargs):
        raise AssertionError("ZipFile ran before the central-directory safety gate")

    monkeypatch.setattr(workbook_module, "ZipFile", unexpected_zip_reader)

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_an_oversized_member_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-member.xlsx")
    with ZipFile(workbook) as archive:
        largest_member = max(member.file_size for member in archive.infolist())
    _append_member(workbook, "xl/media/filler.bin", b"x" * (largest_member + 1))
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_MEMBER_UNCOMPRESSED_BYTES",
        largest_member,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_excessive_aggregate_size_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "oversized-total.xlsx")
    with ZipFile(workbook) as archive:
        original_total = sum(member.file_size for member in archive.infolist())
    _append_member(workbook, "xl/media/filler.bin", b"x")
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_TOTAL_UNCOMPRESSED_BYTES",
        original_total,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_a_compression_bomb_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "compression-bomb.xlsx")
    with ZipFile(workbook) as archive:
        ordinary_ratio = max(
            member.file_size // max(member.compress_size, 1)
            for member in archive.infolist()
        )
    _append_member(workbook, "xl/media/compressed.bin", b"x" * 200_000)
    with ZipFile(workbook) as archive:
        compressed_member = archive.getinfo("xl/media/compressed.bin")
    assert compressed_member.file_size > compressed_member.compress_size * ordinary_ratio
    monkeypatch.setattr(
        workbook_module,
        "_OOXML_ARCHIVE_MAX_COMPRESSION_RATIO",
        ordinary_ratio,
    )

    _reject_before_workbook_readers(monkeypatch, workbook)


@pytest.mark.parametrize(
    ("member_name", "expected_detail"),
    [
        ("../not-a-workbook-part.xml", "unsafe or non-canonical"),
        ("XL/WORKBOOK.XML", "ambiguous"),
    ],
)
def test_archive_preflight_rejects_ambiguous_or_unsafe_member_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    member_name: str,
    expected_detail: str,
) -> None:
    workbook = make_model(tmp_path / "unsafe-member-path.xlsx")
    _append_member(workbook, member_name, b"untrusted")

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert expected_detail in message
    assert member_name not in message


def test_archive_preflight_rejects_duplicate_zip_members_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "duplicate-member.xlsx")
    with ZipFile(workbook) as archive:
        workbook_payload = archive.read("xl/workbook.xml")
    with pytest.warns(UserWarning, match="Duplicate name"):
        _append_member(workbook, "xl/workbook.xml", workbook_payload)

    message = _reject_before_workbook_readers(monkeypatch, workbook)

    assert "ambiguous" in message
    assert "workbook.xml" not in message


def test_archive_preflight_rejects_encrypted_central_member_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "encrypted-member.xlsx")
    contents = bytearray(workbook.read_bytes())
    central_directory_offset = _last_central_directory_offset(contents)
    flag_offset = central_directory_offset + 8
    flag_bits = int.from_bytes(contents[flag_offset : flag_offset + 2], "little")
    contents[flag_offset : flag_offset + 2] = (flag_bits | 0x1).to_bytes(2, "little")
    workbook.write_bytes(contents)

    assert "encrypted" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_unicode_path_aliases_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "unicode-path-alias.xlsx")
    original_name = "xl/media/original.bin"
    unicode_alias = "xl/media/alias.bin"
    entry = ZipInfo(original_name)
    entry.extra = struct.pack(
        "<HHBI",
        0x7075,
        5 + len(unicode_alias.encode("utf-8")),
        1,
        binascii.crc32(original_name.encode("ascii")),
    ) + unicode_alias.encode("utf-8")
    with ZipFile(workbook, "a") as archive:
        archive.writestr(entry, b"untrusted")

    assert "Unicode-path aliases" in _reject_before_workbook_readers(
        monkeypatch,
        workbook,
    )


def test_archive_preflight_rejects_malformed_zip64_member_metadata_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "malformed-zip64.xlsx")
    contents = bytearray(workbook.read_bytes())
    central_directory_offset = _last_central_directory_offset(contents)
    uncompressed_size_offset = central_directory_offset + 24
    contents[uncompressed_size_offset : uncompressed_size_offset + 4] = b"\xff" * 4
    workbook.write_bytes(contents)

    assert "ZIP64" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_archive_preflight_rejects_local_header_mismatch_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "local-header-mismatch.xlsx")
    contents = bytearray(workbook.read_bytes())
    local_flag_offset = 6
    flag_bits = int.from_bytes(contents[local_flag_offset : local_flag_offset + 2], "little")
    contents[local_flag_offset : local_flag_offset + 2] = (flag_bits | 0x8).to_bytes(
        2,
        "little",
    )
    workbook.write_bytes(contents)

    assert "local member metadata is inconsistent" in _reject_before_workbook_readers(
        monkeypatch,
        workbook,
    )


def test_archive_preflight_rejects_symbolic_link_members_before_readers(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = make_model(tmp_path / "symbolic-link.xlsx")
    link = ZipInfo("xl/media/link")
    link.create_system = 3
    link.external_attr = (stat.S_IFLNK | 0o777) << 16
    with ZipFile(workbook, "a") as archive:
        archive.writestr(link, b"not a real part")

    assert "symbolic-link" in _reject_before_workbook_readers(monkeypatch, workbook)


def test_cli_surfaces_archive_preflight_as_an_input_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    workbook = make_model(tmp_path / "cli-limit.xlsx")
    monkeypatch.setattr(workbook_module, "_OOXML_ARCHIVE_MAX_ENTRY_COUNT", 1)

    assert main(["profile", str(workbook)]) == 2

    assert "safety preflight" in capsys.readouterr().err
