"""Tahap 2: raw JSON MATRIX -> JSON final (satu record per colorway) + antrean needs_review."""
import re
from collections import Counter, defaultdict
from datetime import date

from .errors import Matrix2ImError
from .legend_config import DESCRIPTOR_RE, IM_COLOR_KEYS, SOURCE_NAME, TARGET_NAME
from .review import Code, Problem
from .tag_rules import build_ws_tag_color, unique_yarns
from .whs_rules import color_whs, final_whs

GROUPS = (("color_above_black_line", "above_black"),
          ("color_between_black_and_red_line", "between_black_red"),
          ("color_below_red_line", "below_red"))

# Kode yang bukan berasal dari TAG/font/WHS-capped membawa instruksi TAG hanya bila ada di sini.
_WITH_TAG_INSTRUCTION = frozenset({Code.TAG_CONDITIONAL, Code.TAG_CN_OUT_OF_RANGE, Code.TAG_EMPTY,
                                   Code.FONT_COLOR_NO_RULE, Code.WHS_STYLE_UNKNOWN, Code.WHS_CONFLICT,
                                   Code.TOO_MANY_YARNS})

RULES = {
    "DESCRIPTION": "nama style dari MATRIX (tanpa size); berlaku untuk semua size di IM",
    "COLOR1-COLOR18": "1 benang = 1 kolom, nilai apa adanya dari MATRIX; benang berulang (stripe) ditulis sekali; "
                      "slot tanpa benang tidak dicantumkan (tidak ditulis ke IM); benang yang sama dianggap sama "
                      "tanpa membedakan huruf besar/kecil",
    "WS TAG COLOR": "instruksi TAG, C1/C2/... diganti nama warna benang (tanpa prefix YCA/YST dan kode); "
                    "STRIPED -> STRIPE; alias nama (mis. BLACK B -> BLACK); akhiran dobel (STRIPE STRIPE) "
                    "dijadikan satu; null = belum ada aturan, tidak ditulis",
    "WHS": "kolom CE. Warna: di bawah garis merah = SY; di atas garis merah = 000 & SY, "
           "kecuali font ungu/pink = 000. Style (background No): ungu 000, abu 000 & SY, merah SY. "
           "WHS style = batas atas (irisan)",
    "_ref": "asal di MATRIX: sheet (sub-koleksi) No | sel | posisi terhadap garis; tidak ditulis ke IM",
    "needs_review.code": "kode mesin: " + ", ".join(c.value for c in Code),
}


def _item(record, code: Code, desc: str, colorway: str, issue: str, ref) -> dict:
    return {"record": record, "code": code.value, "DESCRIPTION": desc, "colorway": colorway,
            "issue": issue, "_ref": ref}


def _with_extra(item: dict, **extra) -> dict:
    """Sisipkan field opsional sebelum `_ref` supaya urutan key tetap."""
    ref = item.pop("_ref")
    item.update(extra)
    item["_ref"] = ref
    return item


def _build_records(raw: dict) -> tuple[list[dict], list[dict]]:
    """Bangun record dan item review per-record (urutan: TAG, font, WHS, TOO_MANY, DESCRIPTOR)."""
    records, review = [], []
    descriptor = re.compile(DESCRIPTOR_RE)
    for style in raw["styles"]:
        for group_key, group in GROUPS:
            for cw in style[group_key]:
                yarns = unique_yarns(cw["yarns"])
                rec = {"DESCRIPTION": style["style"]}
                for key, yarn in zip(IM_COLOR_KEYS, yarns):
                    rec[key] = yarn
                tag, p_tag = build_ws_tag_color(style["tag"], yarns)
                col, p_font = color_whs(group, cw["font_color"])
                whs, p_whs = final_whs(style.get("style_whs"), col)
                rec["WS TAG COLOR"], rec["WHS"] = tag, whs
                section = f" ({style['section']})" if style["section"] else ""
                ref = f"{style['sheet']}{section} No.{style['no']} | {cw['cell']} | {group}"
                rec["_ref"] = ref
                idx = len(records)
                records.append(rec)

                shown = yarns[:len(IM_COLOR_KEYS)]
                colorway = "/".join(shown)
                problems = [p for p in (p_tag, p_font, p_whs) if p]
                if len(yarns) > len(IM_COLOR_KEYS):
                    problems.append(Problem(Code.TOO_MANY_YARNS,
                                            f"{len(yarns)} benang unik, slot COLOR hanya {len(IM_COLOR_KEYS)}"))
                for p in problems:
                    item = _item(idx, p.code, rec["DESCRIPTION"], colorway, p.issue, ref)
                    if p.code in _WITH_TAG_INSTRUCTION:
                        _with_extra(item, tag_instruction=style["tag"])
                    review.append(item)
                desc = [y for y in shown if descriptor.search(y.upper())]
                if desc:
                    item = _item(idx, Code.COLOR_IS_DESCRIPTOR, rec["DESCRIPTION"], colorway,
                                 f"COLOR berisi deskripsi, bukan nama benang: {desc}", ref)
                    review.append(_with_extra(item, comment_di_matrix=cw.get("comment")))
    return records, review


def _duplicate_items(records: list[dict]) -> list[dict]:
    """Style+colorway yang muncul >1 kali (biasanya di 2 sheet) -> satu item per kelompok."""
    seen: dict[tuple[str, str], list[int]] = defaultdict(list)
    for i, rec in enumerate(records):
        colorway = "/".join(rec[k] for k in IM_COLOR_KEYS if k in rec)
        seen[(rec["DESCRIPTION"], colorway)].append(i)
    items = []
    for (desc, colorway), idx in seen.items():
        if len(idx) < 2:
            continue
        distinct = {records[j]["WHS"] for j in idx}
        # None diurutkan paling akhir agar tidak crash membandingkan None dengan str.
        whs_values = sorted(distinct, key=lambda v: (v is None, v or ""))
        issue = "style muncul di 2 sheet MATRIX, record dobel"
        if len(distinct) > 1:
            issue += f" dan WHS beda: {whs_values}"
        item = _item(idx, Code.DUPLICATE_STYLE_ACROSS_SHEETS, desc, colorway, issue,
                     [records[j]["_ref"] for j in idx])
        items.append(_with_extra(item, whs_values=whs_values))
    return items


def _sort_key(item: dict) -> int:
    rec = item["record"]
    if isinstance(rec, int):
        return rec
    if isinstance(rec, list):
        return rec[0]
    return 10**6


_STYLE_KEYS = ("no", "style", "tag", "sheet", "section")
_COLORWAY_KEYS = ("yarns", "font_color", "cell")


def _check_structure(raw: dict) -> None:
    """Tolak raw JSON yang kuncinya tidak lengkap dengan pesan jelas, bukan KeyError mentah.

    Raw JSON bisa diedit tangan atau berasal dari versi lain; lebih baik gagal di awal
    dengan nama key dan nomor style daripada di tengah transformasi.
    """
    if not isinstance(raw.get("warnings", []), list):
        raise Matrix2ImError("Raw JSON tidak valid: 'warnings' harus berupa list")
    for i, st in enumerate(raw["styles"], 1):
        if not isinstance(st, dict):
            raise Matrix2ImError(f"Raw JSON tidak valid: style #{i} bukan objek")
        for key in _STYLE_KEYS:
            if key not in st:
                raise Matrix2ImError(f"Raw JSON tidak valid: style #{i} tidak punya key '{key}'")
        for group_key, _ in GROUPS:
            if not isinstance(st.get(group_key), list):
                raise Matrix2ImError(
                    f"Raw JSON tidak valid: style #{i} key '{group_key}' tidak ada atau bukan list")
            for j, cw in enumerate(st[group_key], 1):
                if not isinstance(cw, dict):
                    raise Matrix2ImError(
                        f"Raw JSON tidak valid: style #{i} '{group_key}' colorway #{j} bukan objek")
                for key in _COLORWAY_KEYS:
                    if key not in cw:
                        raise Matrix2ImError(f"Raw JSON tidak valid: style #{i} '{group_key}' "
                                             f"colorway #{j} tidak punya key '{key}'")
                if not isinstance(cw["yarns"], list):
                    raise Matrix2ImError(f"Raw JSON tidak valid: style #{i} '{group_key}' "
                                         f"colorway #{j} key 'yarns' bukan list")


def transform(raw: dict, generated: str | None = None) -> dict:
    """Ubah raw JSON MATRIX jadi dokumen final: records, needs_review, summary."""
    if not isinstance(raw, dict) or not isinstance(raw.get("styles"), list):
        raise Matrix2ImError("Bukan raw JSON MATRIX: key 'styles' tidak ada")
    _check_structure(raw)
    records, review = _build_records(raw)
    review += _duplicate_items(records)
    for w in raw.get("warnings", []):
        review.append({"record": None, "code": Code.EXTRACTION_WARNING.value,
                       "issue": f"warning ekstraksi: {w}"})
    review.sort(key=_sort_key)
    summary = {
        "total_records": len(records),
        "total_styles": len({r["DESCRIPTION"] for r in records}),
        "WHS": dict(Counter(r["WHS"] for r in records)),
        "WS_TAG_COLOR_filled": sum(1 for r in records if r["WS TAG COLOR"]),
        "WS_TAG_COLOR_null": sum(1 for r in records if not r["WS TAG COLOR"]),
        "needs_review_items": len(review),
        "needs_review_by_issue": dict(Counter(x["code"] for x in review)),
    }
    return {"target": TARGET_NAME, "source": SOURCE_NAME,
            "generated": generated or date.today().isoformat(),
            "rules": RULES, "summary": summary, "records": records, "needs_review": review}
