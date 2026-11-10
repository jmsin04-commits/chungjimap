"""신지민_데이터.xlsx → data.json

엑셀을 고친 뒤 이 스크립트를 다시 실행하면 app.html이 읽는 data.json이 새로 만들어집니다.
    python make_data.py
필요한 패키지: openpyxl (pip install openpyxl)
"""
import json
import re
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
SRC = HERE / "신지민_데이터.xlsx"
OUT = HERE / "data.json"

UNKNOWN = {"", "확인안됨", "없음", None}


def clean(v):
    if v is None:
        return ""
    return str(v).strip()


def first_link(v):
    m = re.search(r"https?://\S+", clean(v))
    return m.group(0) if m else ""


def lines(v):
    v = clean(v)
    if v in UNKNOWN:
        return []
    return [x.strip() for x in v.split("\n") if x.strip()]


def load_sheet(ws, kind):
    rows = list(ws.iter_rows(values_only=True))
    head = [clean(h) for h in rows[0]]
    col = {h: i for i, h in enumerate(head)}

    def get(r, name):
        i = col.get(name)
        return r[i] if i is not None and i < len(r) else None

    places = []
    for r in rows[1:]:
        name = clean(get(r, "이름"))
        if not name:
            continue
        tags = [t.strip() for t in clean(get(r, "기타 특징")).split(",")]
        tags = [t for t in tags if t and t not in UNKNOWN]
        region = clean(get(r, "지역"))
        address = clean(get(r, "위치"))
        places.append({
            "kind": kind,
            "name": name,
            "address": "" if address in UNKNOWN else address,
            "map": first_link(get(r, "지도")),
            "checked": clean(get(r, "확인된 날"))[:10],
            "recs": int(get(r, "추천 수") or 1),
            "region": "" if region in UNKNOWN else region,
            "hours": lines(get(r, "영업시간")),
            "price": lines(get(r, "가격")) if kind == "맛집" else [],
            "tags": tags,
        })
    places.sort(key=lambda p: (-p["recs"], p["name"]))
    return places


def main():
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    data = {"맛집": load_sheet(wb["맛집"], "맛집"),
            "숨은장소": load_sheet(wb["숨은장소"], "숨은장소")}
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"data.json: 맛집 {len(data['맛집'])}곳, 숨은장소 {len(data['숨은장소'])}곳")


if __name__ == "__main__":
    main()
