import requests
import pandas as pd
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import time
import urllib.parse


SERVICE_KEY = ""

# API 기본 URL
BASE_URL = "http://apis.data.go.kr/1230000/ad/BidPublicInfoService"

# 조회할 4가지 오퍼레이션 목록
OPERATIONS = {
    "공사": "getBidPblancListInfoCnstwkPPSSrch",
    "용역": "getBidPblancListInfoServcPPSSrch",
    "외자": "getBidPblancListInfoFrgcptPPSSrch",
    "물품": "getBidPblancListInfoThngPPSSrch"
}


class G2BEmailCollector:
    def __init__(self, service_key):
        self.service_key = service_key

    def get_date_chunks(self, days=60):
        """
        API 부하를 줄이고 데이터 누락을 방지하기 위해
        전체 기간을 30일 단위로 쪼개서 리스트로 반환
        """
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)

        chunks = []
        current_start = start_date

        while current_start < end_date:
            current_end = current_start + timedelta(days=29)
            if current_end > end_date:
                current_end = end_date

            # API 포맷: YYYYMMDDHHMM
            str_start = current_start.strftime('%Y%m%d%H%M')
            str_end = current_end.strftime('%Y%m%d%H%M')

            chunks.append((str_start, str_end))
            current_start = current_end + timedelta(minutes=1)  # 1분 뒤부터 다시 시작

        return chunks

    def fetch_data(self, operation_name, operation_code, start_dt, end_dt):
        """특정 오퍼레이션에 대해 API 호출 및 데이터 파싱"""
        url = f"{BASE_URL}/{operation_code}"
        all_rows = []
        page_no = 1

        while True:
            params = {
                'ServiceKey': self.service_key,
                'numOfRows': '900',  # 한 번에 최대 조회 수
                'pageNo': str(page_no),
                'inqryDiv': '1',  # 1: 공고게시일시 기준
                'inqryBgnDt': start_dt,
                'inqryEndDt': end_dt,
                'type': 'xml'  # XML 포맷 명시
            }

            try:
                response = requests.get(url, params=params, timeout=30)

                if response.status_code != 200:
                    print(f"[{operation_name}] HTTP 에러: {response.status_code}")
                    break

                root = ET.fromstring(response.content)
                result_msg = root.find('.//resultMsg').text if root.find('.//resultMsg') is not None else ""

                # 데이터 없음 처리
                if "NO DATA" in result_msg.upper() or "조회된 데이터가 없습니다" in result_msg:
                    break

                items = root.findall('.//item')
                if not items:
                    break

                for item in items:
                    # 이메일 추출
                    email = self._get_text(item, 'ntceInsttOfclEmailAdrs')

                    # 이메일이 있는 경우에만 데이터 수집
                    if email and '@' in email:
                        row = {
                            '분야': operation_name,
                            '공고번호': self._get_text(item, 'bidNtceNo'),
                            '공고명': self._get_text(item, 'bidNtceNm'),
                            '공고기관': self._get_text(item, 'ntceInsttNm'),
                            '담당자명': self._get_text(item, 'ntceInsttOfclNm'),
                            '전화번호': self._get_text(item, 'ntceInsttOfclTelNo'),
                            '이메일': email,
                            '공고일시': self._get_text(item, 'bidNtceDt')
                        }
                        all_rows.append(row)

                print(f"[{operation_name}] {start_dt[:8]}~{end_dt[:8]} - {page_no}페이지 수집 완료 ({len(items)}건 조회)")

                # 페이징 탈출 조건 (조회된 데이터가 요청한 row 수보다 적으면 마지막 페이지)
                if len(items) < int(params['numOfRows']):
                    break

                page_no += 1
                time.sleep(0.2)  # API 서버 부하 방지

            except Exception as e:
                print(f"[{operation_name}] 오류 발생: {e}")
                break

        return all_rows

    def _get_text(self, item, tag):
        element = item.find(tag)
        return element.text.strip() if element is not None and element.text else ""


def main():
    print("나라장터 이메일 수집기를 시작합니다... (최근 2개월 기준)")

    collector = G2BEmailCollector(SERVICE_KEY)
    date_chunks = collector.get_date_chunks(days=60)  # 최근 2개월

    total_data = []

    # 1. 날짜 구간별 순회
    for start_dt, end_dt in date_chunks:
        # 2. 4가지 업무 분야(공사, 용역, 외자, 물품) 순회
        for op_name, op_code in OPERATIONS.items():
            rows = collector.fetch_data(op_name, op_code, start_dt, end_dt)
            total_data.extend(rows)

    if not total_data:
        print("수집된 데이터가 없습니다.")
        return

    # 3. 데이터프레임 변환
    df = pd.DataFrame(total_data)

    print(f"\n총 수집된 건수: {len(df)}건")

    # 4. 중복 제거 (이메일 기준)
    # 공고번호나 담당자가 다르더라도, 같은 이메일이면 중복으로 간주하고 제거 (가장 최근 공고 기준 남김)
    df_unique = df.drop_duplicates(subset=['이메일'], keep='first')

    print(f"중복(이메일 기준) 제거 후 건수: {len(df_unique)}건")

    # 5. 엑셀 저장
    today_str = datetime.now().strftime('%Y-%m-%d')
    file_name = f"나라장터_담당자이메일_{today_str}.xlsx"

    try:
        df_unique.to_excel(file_name, index=False)
        print(f"\n[성공] '{file_name}' 파일로 저장되었습니다.")
    except Exception as e:
        print(f"\n[오류] 엑셀 저장 실패: {e}")


if __name__ == "__main__":
    main()
