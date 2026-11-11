"""styles.xml: kunci warna latar dan penambahan format baru tanpa mengubah format lama."""
import pytest

from matrix_to_im.errors import Matrix2ImError
from matrix_to_im.xlsx_styles import StyleSheet, check_appended, fill_key

_STYLES = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b'<fonts count="1"><font><sz val="11"/></font></fonts>'
    b'<fills count="3"><fill><patternFill patternType="none"/></fill>'
    b'<fill><patternFill patternType="gray125"/></fill>'
    b'<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC99"/><bgColor indexed="64"/></patternFill></fill>'
    b'</fills>'
    b'<borders count="1"><border/></borders>'
    b'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    b'<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    b'<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'
    b'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
    b'<alignment wrapText="1"/></xf></cellXfs>'
    b'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    b'</styleSheet>'
)


@pytest.mark.parametrize("xml, key", [
    (b'<fill><patternFill patternType="none"/></fill>', None),
    (b"<fill><patternFill/></fill>", None),
    (b"<fill/>", None),
    (b'<fill><patternFill patternType="solid"><fgColor rgb="ffff99ff"/></patternFill></fill>', "FF99FF"),
    (b'<fill><patternFill patternType="solid"><fgColor rgb="00969696"/><bgColor indexed="64"/></patternFill></fill>',
     "969696"),
    (b'<fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.15"/></patternFill></fill>',
     "solid theme=0 tint=-0.15"),
    (b'<fill><patternFill patternType="gray125"/></fill>', "gray125"),
])
def test_fill_key(xml, key):
    assert fill_key(xml) == key


def test_xf_fill_key_reads_through_cellxfs():
    st = StyleSheet(_STYLES)
    assert [st.xf_fill_key(i) for i in range(3)] == [None, "FFCC99", None]


def test_xf_index_out_of_range_raises():
    with pytest.raises(Matrix2ImError, match="format sel 9"):
        StyleSheet(_STYLES).xf_fill_key(9)


def test_restyle_reuses_existing_fill_and_xf():
    st = StyleSheet(_STYLES)
    assert st.restyle(0, "FFCC99") == 1          # xf 0 + fill 2 + applyFill = persis xf 1
    assert st.restyle(1, "FFCC99") == 1          # sudah berwarna itu
    assert st.restyle(1, None) == 0              # xf 1 tanpa warna = xf 0
    assert st.restyle(2, "FFCC99") == 3          # xf 2 (wrap text) + FFCC99: format baru
    assert st.restyle(2, "FFCC99") == 3          # cache
    out = st.to_bytes()
    assert b'<fills count="3">' in out           # fill FFCC99 dan fill kosong sudah ada
    assert b'<cellXfs count="4">' in out
    check_appended(_STYLES, out)


def test_restyle_appends_solid_fill_and_keeps_other_attributes():
    st = StyleSheet(_STYLES)
    idx = st.restyle(2, "FF99FF")
    assert idx == 3
    new = StyleSheet(st.to_bytes())
    assert new.xf_fill_key(3) == "FF99FF"
    assert b'<fill><patternFill patternType="solid"><fgColor rgb="FFFF99FF"/><bgColor indexed="64"/>' \
           b'</patternFill></fill></fills>' in st.to_bytes()
    xf = new.xfs[3]
    assert b'applyAlignment="1"' in xf and b'<alignment wrapText="1"/>' in xf
    assert b'fillId="3"' in xf and b'applyFill="1"' in xf
    assert b'<fills count="4">' in st.to_bytes()


def test_restyle_reuses_identical_xf_from_earlier_run():
    first = StyleSheet(_STYLES)
    first.restyle(0, "FF99FF")
    again = StyleSheet(first.to_bytes())
    assert again.restyle(0, "FF99FF") == 3       # xf hasil run sebelumnya dipakai ulang, tidak menumpuk
    assert again.to_bytes() == first.to_bytes()


def test_no_change_returns_same_bytes():
    st = StyleSheet(_STYLES)
    st.restyle(1, "FFCC99")
    assert st.to_bytes() == _STYLES


def test_check_appended_rejects_changed_old_entries():
    st = StyleSheet(_STYLES)
    st.restyle(0, "FF99FF")
    out = st.to_bytes()
    with pytest.raises(Matrix2ImError, match="styles.xml"):
        check_appended(_STYLES, out.replace(b'fillId="2" borderId="0" xfId="0" applyFill="1"',
                                            b'fillId="1" borderId="0" xfId="0" applyFill="1"'))
    with pytest.raises(Matrix2ImError, match="styles.xml"):
        check_appended(_STYLES, out.replace(b'<sz val="11"/>', b'<sz val="12"/>'))
    with pytest.raises(Matrix2ImError, match="styles.xml"):
        check_appended(_STYLES, out.replace(b'<fills count="4">', b'<fills count="9">'))


def test_missing_sections_raise():
    with pytest.raises(Matrix2ImError, match="fills"):
        StyleSheet(b"<styleSheet><cellXfs count=\"0\"></cellXfs></styleSheet>")
