"""엑셀의 주소(위치 열)를 좌표로 바꿔 '위도'·'경도' 열에 채웁니다. (네이버 클라우드 Geocoding)

준비
  1. 네이버 클라우드 콘솔 → Maps → Application 등록에서 'Geocoding'과 'Dynamic Map'을 체크하고
     Web 서비스 URL에 https://jmsin04-commits.github.io 를 넣습니다.
  2. 발급된 Client ID / Client Secret을 환경 변수로 넣고 실행합니다.
       (맥·리눅스)  NCP_KEY_ID=아이디 NCP_KEY=시크릿 python geocode.py
       (윈도우)     set NCP_KEY_ID=아이디 & set NCP_KEY=시크릿 & python geocode.py
     시크릿은 저장소에 올리지 마세요. 앱(app.html)에는 Client ID만 들어갑니다.

동작
  - 맛집 시트에서 위도·경도가 비어 있고 주소가 있는 줄만 변환합니다. (--all 이면 숨은장소 시트도)
  - 이미 채워진 칸은 건드리지 않으니 여러 번 돌려도 됩니다.
  - 실행 전에 엑셀을 '신지민_데이터.backup.xlsx'로 복사해 둡니다.
  - 끝나면 못 찾은 곳 목록을 보여줍니다. 그 곳은 지도에서 직접 찍어 위도·경도를 적으면 됩니다.
  - 그다음 python make_data.py 를 실행하면 앱에 반영됩니다.
"""
import json
import os
import shutil
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
SRC = HERE / "신지민_데이터.xlsx"
BACKUP = HERE / "신지민_데이터.backup.xlsx"
URL = "https://maps.apigw.ntruss.com/map-geocode/v2/geocode?query="


def geocode(address, key_id, key):
    req = urllib.request.Request(URL + urllib.parse.quote(address), headers={
        "x-ncp-apigw-api-key-id": key_id,
        "x-ncp-apigw-api-key": key,
        "Accept": "application/json",
    })
    with urllib.request.urlopen(req, timeout=10) as res:
        body = json.load(res)
    hits = body.get("addresses") or []
    if not hits:
        return None
    return float(hits[0]["y"]), float(hits[0]["x"])  # (위도, 경도)


def column(ws, name):
    head = [c.value for c in ws[1]]
    if name in head:
        return head.index(name) + 1
    col = ws.max_column + 1
    ws.cell(row=1, column=col, value=name)
    return col


def main():
    key_id, key = os.environ.get("NCP_KEY_ID"), os.environ.get("NCP_KEY")
    if not key_id or not key:
        sys.exit("NCP_KEY_ID와 NCP_KEY 환경 변수를 먼저 넣어주세요. (파일 맨 위 설명 참고)")
    sheets = ["맛집", "숨은장소"] if "--all" in sys.argv else ["맛집"]

    shutil.copyfile(SRC, BACKUP)
    wb = openpyxl.load_workbook(SRC)
    done, missed = 0, []
    for name in sheets:
        ws = wb[name]
        head = [c.value for c in ws[1]]
        c_name, c_addr = head.index("이름") + 1, head.index("위치") + 1
        c_lat, c_lng = column(ws, "위도"), column(ws, "경도")
        for row in range(2, ws.max_row + 1):
            place = ws.cell(row, c_name).value
            addr = str(ws.cell(row, c_addr).value or "").strip()
            if not place or ws.cell(row, c_lat).value:
                continue
            if not addr or addr in ("확인안됨", "없음"):
                missed.append((name, place, "주소 없음"))
                continue
            try:
                hit = geocode(addr, key_id, key)
            except Exception as e:  # 네트워크·키 오류
                sys.exit(f"요청 실패: {e}\n키와 콘솔의 Geocoding 체크를 확인하세요. 지금까지 결과는 저장하지 않았습니다.")
            if hit:
                ws.cell(row, c_lat, round(hit[0], 6))
                ws.cell(row, c_lng, round(hit[1], 6))
                done += 1
            else:
                missed.append((name, place, addr))
            time.sleep(0.1)
    wb.save(SRC)
    print(f"좌표를 채운 곳: {done}곳 (백업: {BACKUP.name})")
    if missed:
        print(f"못 찾은 곳 {len(missed)}곳 — 지도에서 직접 찍어 위도·경도를 적어주세요:")
        for n, p, a in missed:
            print(f"  [{n}] {p} — {a}")
    print("다음: python make_data.py")


if __name__ == "__main__":
    main()
