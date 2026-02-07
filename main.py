import requests
import pandas as pd
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import os
import time

# === [설정] 서비스 키 입력 ===
SERVICE_KEY = ""


class G2BAPIClient:
    def __init__(self, service_key):
        self.base_url = "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch"
        self.service_key = service_key

    def fetch_bid_notices(self, search_params):
        """API 1회 호출"""
        params = {
            'ServiceKey': self.service_key,
            'numOfRows': search_params.get('numOfRows', 100),
            'pageNo': search_params.get('pageNo', 1),
            'inqryDiv': search_params.get('inqryDiv', '1'),
            'inqryBgnDt': search_params['inqryBgnDt'],
            'inqryEndDt': search_params['inqryEndDt']
        }

        if 'bidNtceNm' in search_params and search_params['bidNtceNm']:
            params['bidNtceNm'] = search_params['bidNtceNm']

        try:
            response = requests.get(self.base_url, params=params, timeout=30)
            if response.status_code != 200:
                print(f"  [HTTP 오류] {response.status_code}")
                return []

            root = ET.fromstring(response.content)
            result_code = root.find('.//resultCode').text if root.find('.//resultCode') is not None else None

            if result_code != '00':
                result_msg = root.find('.//resultMsg').text if root.find('.//resultMsg') is not None else "알 수 없는 오류"
                # 데이터가 없는 경우(조회 결과 없음)는 오류가 아님
                if "조회된 데이터가 없습니다" in result_msg:
                    return []
                print(f"  [API 메시지] {result_msg}")
                return []

            items = root.findall('.//item')
            return self._parse_items(items)

        except Exception as e:
            print(f"  [시스템 오류] {e}")
            return []

    def _parse_items(self, items):
        result = []
        for item in items:
            data = {
                'bidNtceNo': self._get_text(item, 'bidNtceNo'),
                'rgstTyNm': self._get_text(item, 'rgstTyNm'),
                'ntceKindNm': self._get_text(item, 'ntceKindNm'),
                'bidNtceDt': self._get_text(item, 'bidNtceDt'),
                'bidNtceNm': self._get_text(item, 'bidNtceNm'),
                'ntceInsttNm': self._get_text(item, 'ntceInsttNm'),
                'dminsttNm': self._get_text(item, 'dminsttNm'),
                'ntceInsttOfclNm': self._get_text(item, 'ntceInsttOfclNm'),
                'ntceInsttOfclTelNo': self._get_text(item, 'ntceInsttOfclTelNo'),
                'bidClseDt': self._get_text(item, 'bidClseDt'),
                'opengDt': self._get_text(item, 'opengDt')
            }
            result.append(data)
        return result

    def _get_text(self, item, tag_name):
        element = item.find(tag_name)
        return element.text if element is not None and element.text else ''

    def fetch_all_pages(self, search_params):
        """페이징 처리"""
        all_data = []
        page_no = 1
        num_of_rows = search_params.get('numOfRows', 100)

        while True:
            search_params['pageNo'] = page_no
            data = self.fetch_bid_notices(search_params)

            if not data:
                break

            all_data.extend(data)
            if len(data) < num_of_rows:
                break
            page_no += 1
            time.sleep(0.1)  # API 부하 방지용 짧은 대기

        return all_data


def get_automatic_date_ranges():
    """
    [수정됨] 오늘 기준으로 최근 2주간의 날짜 범위를 생성
    """
    now = datetime.now()

    # 시작일: 오늘로부터 2주(14일) 전
    start_date = (now - timedelta(weeks=2)).replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = now

    ranges = []
    current_start = start_date

    # 2주 기간이라도 월이 바뀔 수 있으므로 30일 단위 끊기 로직은 유지 (안전장치)
    while current_start < end_date:
        current_end = current_start + timedelta(days=29)

        if current_end > end_date:
            current_end = end_date

        str_start = current_start.strftime('%Y%m%d%H%M')
        str_end = current_end.strftime('%Y%m%d%H%M')

        ranges.append((str_start, str_end))
        current_start = current_end + timedelta(minutes=1)

    return ranges


def save_to_excel(df):
    if df.empty:
        print("저장할 데이터가 없습니다.")
        return

    col_map = {
        'bidNtceNo': '공고번호',
        'rgstTyNm': '등록유형',
        'ntceKindNm': '공고종류',
        'bidNtceDt': '공고일시',
        'bidNtceNm': '공고명',
        'ntceInsttNm': '공고기관',
        'dminsttNm': '수요기관',
        'ntceInsttOfclNm': '담당자',
        'ntceInsttOfclTelNo': '전화번호',
        'bidClseDt': '마감일시',
        'opengDt': '개찰일시'
    }

    # 필요한 컬럼만 선택 및 이름 변경
    save_df = df[list(col_map.keys())].rename(columns=col_map)

    # 엑셀 파일명 생성 (오늘 날짜 포함)
    today_str = datetime.now().strftime('%Y-%m-%d')
    filename = f'IT장비_입찰공고_최근2주({today_str}).xlsx'

    try:
        save_df.to_excel(filename, index=False)
        print(f"\n[성공] 엑셀 저장 완료: {filename}")
        print(f"총 공고 수: {len(save_df)}건")
    except PermissionError:
        print(f"\n[오류] '{filename}' 파일이 이미 열려있습니다. 파일을 닫고 다시 실행해주세요.")
    except Exception as e:
        print(f"\n[오류] 엑셀 저장 중 문제가 발생했습니다: {e}")


def main():
    client = G2BAPIClient(SERVICE_KEY)
    # 검색하고 싶은 키워드 리스트
    target_keywords = ["서버", "GPU", "렌탈", "워크스테이션", "임대", "RISE", "혁신"]

    # 2주 기간 자동 설정
    date_ranges = get_automatic_date_ranges()

    print(f"조회 기간: {date_ranges[0][0]} ~ {date_ranges[-1][1]} (최근 2주)")
    print("데이터 수집을 시작합니다...")

    all_results = []

    for keyword in target_keywords:
        print(f"\n--- '{keyword}' 검색 중 ---")
        for start_dt, end_dt in date_ranges:
            params = {
                'inqryDiv': '1',
                'inqryBgnDt': start_dt,
                'inqryEndDt': end_dt,
                'numOfRows': 100,
                'bidNtceNm': keyword
            }
            results = client.fetch_all_pages(params)
            if results:
                all_results.extend(results)
                print(f"  > {len(results)}건 발견")
            else:
                print(f"  > 데이터 없음")

    if all_results:
        df = pd.DataFrame(all_results)
        # 공고번호 기준 중복 제거
        df_unique = df.drop_duplicates(subset=['bidNtceNo'])
        print(f"\n[최종 집계] 중복 제거 후 총 {len(df_unique)}건")
        save_to_excel(df_unique)
    else:
        print("\n최근 2주간 해당 키워드로 조회된 공고가 없습니다.")


if __name__ == "__main__":
    main()