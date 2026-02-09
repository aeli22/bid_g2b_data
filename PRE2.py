import requests
import pandas as pd
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import os
import time

# === [설정] 서비스 키 입력 ===
SERVICE_KEY = ""

class G2BPublicRangeClient:
    def __init__(self, service_key):
        # 사전규격정보서비스 베이스 URL [cite: 14]
        self.base_url = "http://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/"
        self.service_key = service_key

        # 분야별 오퍼레이션 명칭 정의 (검색조건 포함) [cite: 16]
        self.operations = {
            '물품': 'getPublicPrcureThngInfoThngPPSSrch',
            '외자': 'getPublicPrcureThngInfoFrgcptPPSSrch',
            '용역': 'getPublicPrcureThngInfoServcPPSSrch',
            '공사': 'getPublicPrcureThngInfoCnstwkPPSSrch'
        }

    def fetch_pre_specs(self, biz_type, search_params):
        """특정 분야(물품/외자/용역/공사) API 1회 호출"""
        op_name = self.operations.get(biz_type)
        if not op_name:
            return []

        url = f"{self.base_url}{op_name}"
        params = {
            'ServiceKey': self.service_key,
            'numOfRows': search_params.get('numOfRows', 100),
            'pageNo': search_params.get('pageNo', 1),
            'inqryDiv': '1',  # 1: 접수일시 기준 [cite: 140]
            'inqryBgnDt': search_params['inqryBgnDt'],
            'inqryEndDt': search_params['inqryEndDt'],
            'prdctClsfcNoNm': search_params.get('keyword', '')  # 품명/사업명 검색 [cite: 140]
        }

        try:
            response = requests.get(url, params=params, timeout=30)
            if response.status_code != 200:
                print(f"  [{biz_type} HTTP 오류] {response.status_code}")
                return []

            root = ET.fromstring(response.content)
            result_code = root.find('.//resultCode').text if root.find('.//resultCode') is not None else None

            if result_code != '00':
                result_msg = root.find('.//resultMsg').text if root.find('.//resultMsg') is not None else "알 수 없는 오류"
                if "조회된 데이터가 없습니다" not in result_msg:
                    print(f"  [{biz_type} API 메시지] {result_msg}")
                return []

            items = root.findall('.//item')
            return self._parse_items(items)

        except Exception as e:
            print(f"  [{biz_type} 시스템 오류] {e}")
            return []

    def _parse_items(self, items):
        """API 응답 메시지 명세에 따른 데이터 파싱 [cite: 23, 143]"""
        result = []
        for item in items:
            data = {
                'bfSpecRgstNo': self._get_text(item, 'bfSpecRgstNo'),  # 사전규격등록번호 [cite: 23]
                'bsnsDivNm': self._get_text(item, 'bsnsDivNm'),       # 업무구분명 [cite: 23]
                'refNo': self._get_text(item, 'refNo'),               # 참조번호 [cite: 23]
                'prdctClsfcNoNm': self._get_text(item, 'prdctClsfcNoNm'), # 품명(사업명) [cite: 23]
                'orderInsttNm': self._get_text(item, 'orderInsttNm'),   # 발주기관명 [cite: 23]
                'rlDminsttNm': self._get_text(item, 'rlDminsttNm'),     # 실수요기관명 [cite: 23]
                'asignBdgtAmt': self._get_text(item, 'asignBdgtAmt'),   # 배정예산금액 [cite: 23]
                'rcptDt': self._get_text(item, 'rcptDt'),               # 접수일시 [cite: 23]
                'opninRgstClseDt': self._get_text(item, 'opninRgstClseDt'), # 의견등록마감일시 [cite: 23]
                'ofclNm': self._get_text(item, 'ofclNm'),               # 담당자명 [cite: 23]
                'ofclTelNo': self._get_text(item, 'ofclTelNo'),         # 담당자전화번호 [cite: 23]
                'swBizObjYn': self._get_text(item, 'swBizObjYn'),       # SW사업대상여부 [cite: 23]
                'dlvrTmlmtDt': self._get_text(item, 'dlvrTmlmtDt'),     # 납품기한일시 [cite: 23]
                'dlvrDaynum': self._get_text(item, 'dlvrDaynum'),       # 납품일수 [cite: 23]
                'rgstDt': self._get_text(item, 'rgstDt'),               # 등록일시 [cite: 23]
                'chgDt': self._get_text(item, 'chgDt'),                 # 변경일시 [cite: 23]
                'bidNtceNoList': self._get_text(item, 'bidNtceNoList'), # 입찰공고번호목록 [cite: 23]
                'prdctDtlList': self._get_text(item, 'prdctDtlList'),   # 물품상세목록 [cite: 23]
                'specDocFileUrl1': self._get_text(item, 'specDocFileUrl1'), # 규격문서파일URL1 [cite: 23]
                'specDocFileUrl2': self._get_text(item, 'specDocFileUrl2'),
                'specDocFileUrl3': self._get_text(item, 'specDocFileUrl3'),
                'specDocFileUrl4': self._get_text(item, 'specDocFileUrl4'),
                'specDocFileUrl5': self._get_text(item, 'specDocFileUrl5')
            }
            result.append(data)
        return result

    def _get_text(self, item, tag_name):
        element = item.find(tag_name)
        return element.text if element is not None and element.text else ''

    def fetch_all_pages(self, biz_type, search_params):
        all_data = []
        page_no = 1
        num_of_rows = search_params.get('numOfRows', 100)

        while True:
            search_params['pageNo'] = page_no
            data = self.fetch_pre_specs(biz_type, search_params)

            if not data:
                break

            all_data.extend(data)
            if len(data) < num_of_rows:
                break
            page_no += 1
            time.sleep(0.2) # API 부하 방지

        return all_data

def get_automatic_date_ranges():
    """오늘 기준으로 최근 1개월(30일)의 날짜 범위를 생성"""
    now = datetime.now()
    # 2주(timedelta(weeks=2))에서 1개월(timedelta(days=30))로 변경
    start_date = (now - timedelta(days=70)).replace(hour=0, minute=0)
    end_date = now

    ranges = []
    str_start = start_date.strftime('%Y%m%d%H%M') # YYYYMMDDHHMM 형식 [cite: 20]
    str_end = end_date.strftime('%Y%m%d%H%M')
    ranges.append((str_start, str_end))
    return ranges

def save_to_excel(df):
    if df.empty:
        print("저장할 데이터가 없습니다.")
        return

    col_map = {
        'bfSpecRgstNo': '사전규격등록번호',
        'bsnsDivNm': '업무구분',
        'refNo': '참조번호',
        'prdctClsfcNoNm': '품명(사업명)',
        'orderInsttNm': '발주기관',
        'rlDminsttNm': '수요기관',
        'asignBdgtAmt': '배정예산',
        'rcptDt': '접수일시',
        'opninRgstClseDt': '의견마감일시',
        'ofclNm': '담당자',
        'ofclTelNo': '전화번호',
        'swBizObjYn': 'SW사업여부',
        'bidNtceNoList': '연관공고번호',
        'rgstDt': '등록일시'
    }

    save_df = df[list(col_map.keys())].rename(columns=col_map)
    today_str = datetime.now().strftime('%Y-%m-%d')
    filename = f'나라장터_사전규격_통합조회_1개월({today_str}).xlsx'

    try:
        save_df.to_excel(filename, index=False)
        print(f"\n[성공] 엑셀 저장 완료: {filename}")
        print(f"총 데이터 수: {len(save_df)}건")
    except Exception as e:
        print(f"\n[오류] 엑셀 저장 중 문제 발생: {e}")


def main():
    client = G2BPublicRangeClient(SERVICE_KEY)
    target_keywords = ["렌탈", "임대", "대여", "임차", "위탁관리"]
    # === [추가] 제외할 키워드 설정 ===
    exclude_keywords = ["차량", "통학버스", "버스"]

    biz_types = ['물품', '외자', '용역', '공사']
    date_ranges = get_automatic_date_ranges()
    start_dt, end_dt = date_ranges[0]

    print(f"조회 기간: {start_dt} ~ {end_dt} (최근 1개월)")
    print("사전규격 데이터 수집을 시작합니다...")

    all_results = []

    for biz in biz_types:
        print(f"\n>>> [{biz}] 분야 검색 시작")
        for keyword in target_keywords:
            params = {
                'inqryBgnDt': start_dt,
                'inqryEndDt': end_dt,
                'numOfRows': 100,
                'keyword': keyword
            }
            results = client.fetch_all_pages(biz, params)
            if results:
                all_results.extend(results)
                print(f"  - '{keyword}': {len(results)}건 발견")
            else:
                print(f"  - '{keyword}': 데이터 없음")

    if all_results:
        df = pd.DataFrame(all_results)
        # 1. 중복 제거
        df_unique = df.drop_duplicates(subset=['bfSpecRgstNo'])

        # 2. [핵심 수정] 제외 키워드 필터링
        # prdctClsfcNoNm(품명/사업명) 컬럼에 제외 키워드가 포함되지 않은 것만 추출
        if not df_unique.empty:
            # 정규표현식 패턴 생성 (차량|통학버스|버스)
            exclude_pattern = '|'.join(exclude_keywords)
            # 해당 패턴을 포함하지 않는(~) 행만 선택
            df_filtered = df_unique[~df_unique['prdctClsfcNoNm'].str.contains(exclude_pattern, case=False, na=False)]

            removed_count = len(df_unique) - len(df_filtered)
            print(f"\n[필터링] 제외 키워드 포함 공고 {removed_count}건 삭제 완료")
            print(f"[최종 집계] {len(df_filtered)}건 저장 예정")

            save_to_excel(df_filtered)
        else:
            print("\n조회된 데이터가 없습니다.")
    else:
        print("\n최근 1개월간 해당 키워드로 조회된 사전규격이 없습니다.")

if __name__ == "__main__":
    main()