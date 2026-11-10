import zipfile

import pytest
from helpers.im_factory import (
    TEST_HASH_SPIN100000,
    TEST_PASSWORD,
    TEST_SALT_B64,
    GroupSpec,
    add_file_sharing,
    build_im_sheet,
)

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.im_lock import (
    NotLocked,
    WrongPassword,
    check_modify_password,
    hash_modify_password,
    owner_file,
    owner_name,
    read_file_sharing,
)


@pytest.fixture
def locked(tmp_path):
    p = build_im_sheet(tmp_path / "F27 IM.xlsx", [GroupSpec("A", [["X"]])])
    return add_file_sharing(p)


def test_hash_matches_known_vector():
    assert hash_modify_password(TEST_PASSWORD, TEST_SALT_B64, 100000) == TEST_HASH_SPIN100000


def test_unsupported_algorithm():
    with pytest.raises(Matrix2ImError, match="tidak didukung"):
        hash_modify_password("x", TEST_SALT_B64, 1, "MD5")


def test_check_password_ok_and_wrong(locked):
    check_modify_password(locked, TEST_PASSWORD)
    with pytest.raises(WrongPassword, match="Password IM MASTER salah"):
        check_modify_password(locked, "salah")


def test_unlocked_file_raises_not_locked(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", [["X"]])])
    assert read_file_sharing(p) is None
    with pytest.raises(NotLocked, match="tidak dikunci"):
        check_modify_password(p, TEST_PASSWORD)


def test_read_file_sharing_fields(locked):
    fs = read_file_sharing(locked)
    assert fs.user_name == "Tester"
    assert fs.spin_count == 1000
    assert fs.read_only_recommended is True
    assert fs.algorithm == "SHA-512"
    assert fs.legacy is False


def test_legacy_reservation_password_rejected(tmp_path):
    p = build_im_sheet(tmp_path / "im.xlsx", [GroupSpec("A", [["X"]])])
    tmp = tmp_path / "legacy.xlsx"
    with zipfile.ZipFile(p) as src, zipfile.ZipFile(tmp, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                xml = data.decode()
                i = xml.find(">", xml.find("<workbook")) + 1
                data = (xml[:i] + '<fileSharing reservationPassword="CC1A"/>' + xml[i:]).encode()
            dst.writestr(item, data)
    assert read_file_sharing(tmp).legacy is True
    with pytest.raises(Matrix2ImError, match="format lama"):
        check_modify_password(tmp, TEST_PASSWORD)


def test_owner_file_full_and_truncated_name(tmp_path):
    im = tmp_path / "F27 IM.xlsx"
    assert owner_file(im) is None
    (tmp_path / "~$lain.xlsx").write_bytes(b"x")
    assert owner_file(im) is None
    (tmp_path / "~$7 IM.xlsx").write_bytes(b"x")
    assert owner_file(im) == tmp_path / "~$7 IM.xlsx"
    short = tmp_path / "IM.xlsx"
    (tmp_path / "~$IM.xlsx").write_bytes(b"x")
    assert owner_file(short) == tmp_path / "~$IM.xlsx"


def test_owner_name_length_prefixed(tmp_path):
    f = tmp_path / "~$a.xlsx"
    f.write_bytes(bytes([5]) + b"Putri" + b"\x00" * 48)
    assert owner_name(f) == "Putri"
    f.write_bytes(b"\xff" * 10)
    assert owner_name(f) == "seseorang"
