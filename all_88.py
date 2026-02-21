import requests
import pandas as pd
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
import time

# === [설정] 서비스 키 입력 ===
SERVICE_KEY = ""

# === 사용자가 요청한 전체 88개 출력 항목 매핑 사전 ===
FIELDS_MAPPING = {
    'bidNtceNo': '입찰공고번호',
    'bidNtceOrd': '입찰공고차수',
    'reNtceYn': '재공고여부',
    'rgstTyNm': '등록유형명',
    'ntceKindNm': '공고종류명',
    'intrbidYn': '국제입찰여부',
    'bidNtceDt': '입찰공고일시',
    'refNo': '참조번호',
    'bidNtceNm': '입찰공고명',
    'ntceInsttCd': '공고기관코드',
    'ntceInsttNm': '공고기관명',
    'dminsttCd': '수요기관코드',
    'dminsttNm': '수요기관명',
    'bidMethdNm': '입찰방식명',
    'cntrctCnclsMthdNm': '계약체결방법명',
    'ntceInsttOfclNm': '공고기관담당자명',
    'ntceInsttOfclTelNo': '공고기관담당자전화번호',
    'ntceInsttOfclEmailAdrs': '공고기관담당자이메일주소',
    'exctvNm': '집행관명',
    'bidQlfctRgstDt': '입찰참가자격등록마감일시',
    'cmmnSpldmdAgrmntRcptdocMethd': '공동수급협정서접수방식',
    'cmmnSpldmdAgrmntClseDt': '공동수급협정마감일시',
    'cmmnSpldmdCorpRgnLmtYn': '공동수급업체지역제한여부',
    'bidBeginDt': '입찰개시일시',
    'bidClseDt': '입찰마감일시',
    'opengDt': '개찰일시',
    'rbidPermsnYn': '재입찰허용여부',
    'pqApplDocRcptMthdNm': 'PQ신청서접수방법명',
    'pqApplDocRcptDt': 'PQ신청서접수일시',
    'arsltApplDocRcptMthdNm': '실적신청서접수방법명',
    'arsltApplDocRcptDt': '실적신청서접수일시',
    'rgnDutyJntcontrctRt': '지역의무공동도급비율',
    'dtlsBidYn': '내역입찰여부',
    'bidPrtcptLmtYn': '입찰참가제한여부',
    'prearngPrceDcsnMthdNm': '예정가격결정방법명',
    'totPrdprcNum': '총예가건수',
    'drwtPrdprcNum': '추첨예가건수',
    'bdgtAmt': '예산금액',
    'presmptPrce': '추정가격',
    'govsplyAmt': '관급금액',
    'aplBssCntnts': '적용기준내용',
    'indstrytyEvlRt': '업종평가비율',
    'mainCnsttyNm': '주공종명',
    'mainCnsttyCnstwkPrearngAmt': '주공종공사예정금액',
    'opengPlce': '개찰장소',
    'dcmtgOprtnDt': '설명회실시일시',
    'dcmtgOprtnPlce': '설명회실시장소',
    'contrctrcnstrtnGovsplyMtrlAmt': '도급자설치관급자재금액',
    'govcnstrtnGovsplyMtrlAmt': '관급자설치관급자재금액',
    'bidPrtcptFee': '입찰참가수수료',
    'bidGrntymnyPaymntYn': '입찰보증금납부여부',
    'crdtrNm': '채권자명',
    'cmmnSpldmdCnum': '공동수급업체수',
    'untyNtceNo': '통합공고번호',
    'cmmnSpldmdMethdCd': '공동수급방식코드',
    'cmmnSpldmdMethdNm': '공동수급방식명',
    'stdNtceDocUrl': '표준공고서URL',
    'brffcBidprcPermsnYn': '지사투찰허용여부',
    'cnsttyAccotShreRateList': '공종별지분율목록',
    'cnstrtnAbltyEvlAmtList': '시공능력평가금액목록',
    'dsgntCmptYn': '지명경쟁여부',
    'arsltCmptYn': '실적경쟁여부',
    'pqEvalYn': 'PQ심사여부',
    'ntceDscrptYn': '공고설명여부',
    'rsrvtnPrceReMkngMthdNm': '예비가격재작성방법명',
    'mainCnsttyPresmptPrce': '주공종추정가격',
    'orderPlanUntyNo': '발주계획통합번호',
    'sucsfbidLwltRate': '낙찰하한율',
    'rgstDt': '등록일시',
    'bfSpecRgstNo': '사전규격등록번호',
    'sucsfbidMthdCd': '낙찰방법코드',
    'sucsfbidMthdNm': '낙찰방법명',
    'chgDt': '변경일시',
    'dminsttOfclEmailAdrs': '수요기관담당자이메일주소',
    'indstrytyLmtYn': '업종제한여부',
    'cnstrtsiteRgnNm': '공사현장지역명',
    'rgnDutyJntcontrctYn': '지역의무공동도급여부',
    'chgNtceRsn': '변경공고사유',
    'rbidOpengDt': '재입찰개찰일시',
    'ciblAplYn': '건설산업법적용대상여부',
    'mtltyAdvcPsblYn': '상호시장진출허용여부',
    'mtltyAdvcPsblYnCnstwkNm': '건설산업법적용대상공사명',
    'VAT': '부가가치세',
    'indutyVAT': '주공종부가가치세',
    'indstrytyMfrcFldEvlYn': '주력분야평가여부',
    'bidWgrnteeRcptClseDt': '입찰보증서접수마감일시',
    'rgnLmtBidLocplcJdgmBssCd': '지역제한입찰소재지판단기준코드',
    'rgnLmtBidLocplcJdgmBssNm': '지역제한입찰소재지판단기준명'
}


class G2BAPIClient:
    def __init__(self, service_key):
        self.base_url = "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
        self.service_key = service_key

    def fetch_bid_notices(self, op_name, biz_type, search_params):
        """API 1회 호출"""
        url = self.base_url + op_name
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
            response = requests.get(url, params=params, timeout=30)
            if response.status_code != 200:
                print(f"    [HTTP 오류] {response.status_code}")
                return []

            root = ET.fromstring(response.content)
            result_code = root.findtext('.//resultCode')

            if result_code != '00':
                result_msg = root.findtext('.//resultMsg', default="알 수 없는 오류")
                if "조회된 데이터가 없습니다" in result_msg:
                    return []
                print(f"    [API 메시지] {result_msg}")
                return []

            items = root.findall('.//item')
            return self._parse_items(items, biz_type)

        except Exception as e:
            print(f"    [시스템 오류] {e}")
            return []

    def _get_text(self, item, tag_name):
        """XML 태그에서 텍스트를 안전하게 추출 (공백 제거)"""
        text = item.findtext(tag_name)
        return text.strip() if text else ''

    def _parse_items(self, items, biz_type):
        """요청된 88개 항목 전체를 반복문으로 동적 추출"""
        result = []
        for item in items:
            data = {'bizType': biz_type}  # 엑셀 첫 번째 열이 될 '업무구분' 추가

            for key in FIELDS_MAPPING.keys():
                val = self._get_text(item, key)

                # 예산금액 누락 방지: 공사는 bdgtAmt, 물품/용역/외자는 asignBdgtAmt를 사용함
                if key == 'bdgtAmt' and not val:
                    val = self._get_text(item, 'asignBdgtAmt')

                # 이메일 누락 방지: 공고기관 이메일이 없으면 수요기관 이메일로 보완
                if key == 'ntceInsttOfclEmailAdrs' and not val:
                    val = self._get_text(item, 'dminsttOfclEmailAdrs')

                data[key] = val

            result.append(data)
        return result

    def fetch_all_pages(self, op_name, biz_type, search_params):
        """페이징 처리"""
        all_data = []
        page_no = 1
        num_of_rows = search_params.get('numOfRows', 100)

        while True:
            search_params['pageNo'] = page_no
            data = self.fetch_bid_notices(op_name, biz_type, search_params)

            if not data:
                break

            all_data.extend(data)
            if len(data) < num_of_rows:
                break
            page_no += 1
            time.sleep(0.1)

        return all_data


def get_user_date_ranges():
    while True:
        try:
            start_str = input("조회 시작일을 입력하세요 (예: 2025-12-12): ").strip()
            end_str = input("조회 마감일을 입력하세요 (예: 2026-01-01): ").strip()

            start_date = datetime.strptime(start_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_str, "%Y-%m-%d").replace(hour=23, minute=59)

            if start_date > end_date:
                print("시작일이 마감일보다 늦을 수 없습니다. 다시 입력해주세요.\n")
                continue
            if start_date > datetime.now():
                print("시작일은 오늘 또는 과거여야 합니다. 다시 입력해주세요.\n")
                continue
            break
        except ValueError:
            print("날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 다시 입력해주세요.\n")

    ranges = []
    current_start = start_date

    while current_start < end_date:
        current_end = current_start + timedelta(days=29, hours=23, minutes=59)
        if current_end > end_date:
            current_end = end_date

        str_start = current_start.strftime('%Y%m%d%H%M')
        str_end = current_end.strftime('%Y%m%d%H%M')

        ranges.append((str_start, str_end))
        current_start = current_end + timedelta(minutes=1)

    return ranges, start_str, end_str


def get_user_keywords():
    while True:
        k_input = input("\n검색할 키워드를 쉼표(,)로 구분하여 입력하세요 (예: 서버, gpu, 워크스테이션): ").strip()
        if not k_input:
            print("키워드를 하나 이상 입력해야 합니다.")
            continue

        keywords = [k.strip() for k in k_input.split(',')]
        keywords = [k for k in keywords if k]

        if keywords:
            return keywords
        else:
            print("유효한 키워드를 입력해주세요.")


def save_to_excel(df, start_str, end_str):
    if df.empty:
        print("저장할 데이터가 없습니다.")
        return

    # 엑셀 출력을 위한 컬럼 이름 매핑 생성 (업무구분 + 요청하신 88개 전체 항목)
    col_map = {'bizType': '업무구분'}
    col_map.update(FIELDS_MAPPING)

    # DataFrame 컬럼 이름 변경
    save_df = df[list(col_map.keys())].rename(columns=col_map)
    filename = f'입찰공고_전분야상세결과_{start_str}_to_{end_str}.xlsx'

    try:
        save_df.to_excel(filename, index=False)
        print(f"\n[성공] '{filename}' 파일로 상세 저장이 완료되었습니다.")
        print(f"총 공고 수: {len(save_df)}건 / 출력 항목 수: {len(save_df.columns)}개")
    except PermissionError:
        print(f"\n[오류] '{filename}' 파일이 이미 열려있습니다. 파일을 닫고 다시 실행해주세요.")
    except Exception as e:
        print(f"\n[오류] 엑셀 저장 중 문제가 발생했습니다: {e}")


def main():
    print("=== 나라장터 전분야(물품/외자/용역/공사) 상세 항목 검색 시스템 ===")

    date_ranges, start_str, end_str = get_user_date_ranges()
    target_keywords = get_user_keywords()

    client = G2BAPIClient(SERVICE_KEY)

    # [cite_start]4가지 업무 분야별 오퍼레이션 명확히 지정 [cite: 18]
    operations = {
        '물품': 'getBidPblancListInfoThngPPSSrch',
        '외자': 'getBidPblancListInfoFrgcptPPSSrch',
        '용역': 'getBidPblancListInfoServcPPSSrch',
        '공사': 'getBidPblancListInfoCnstwkPPSSrch'
    }

    print(f"\n조회 기간: {start_str} ~ {end_str}")
    print(f"검색 키워드: {', '.join(target_keywords)}")
    print("데이터 수집을 시작합니다...\n")

    all_results = []

    for keyword in target_keywords:
        print(f"--- '{keyword}' 키워드 검색 시작 ---")

        for biz_type, op_name in operations.items():
            print(f"  [{biz_type}] 분야 조회 중...")
            keyword_biz_results = []

            for start_dt, end_dt in date_ranges:
                params = {
                    'inqryDiv': '1',
                    'inqryBgnDt': start_dt,
                    'inqryEndDt': end_dt,
                    'numOfRows': 100,
                    'bidNtceNm': keyword
                }

                results = client.fetch_all_pages(op_name, biz_type, params)
                if results:
                    keyword_biz_results.extend(results)

            if keyword_biz_results:
                stripped_target_keyword = keyword.replace(" ", "")
                filtered_results = []

                for item in keyword_biz_results:
                    bid_name = item.get('bidNtceNm', '')
                    if bid_name:
                        stripped_bid_name = bid_name.replace(" ", "")
                        if stripped_target_keyword.lower() in stripped_bid_name.lower():
                            filtered_results.append(item)

                if filtered_results:
                    all_results.extend(filtered_results)
                    print(f"    -> {len(filtered_results)}건 발견")
                else:
                    print(f"    -> 조건에 맞는 데이터 없음")
            else:
                print(f"    -> 데이터 없음")

    if all_results:
        df = pd.DataFrame(all_results)
        df_unique = df.drop_duplicates(subset=['bidNtceNo'])
        print(f"\n[최종 집계] 중복 제거 후 총 {len(df_unique)}건2의 공고가 추출되었습니다.")

        save_to_excel(df_unique, start_str, end_str)
    else:
        print("\n입력하신 조건으로 조회된 공고가 없습니다.")


if __name__ == "__main__":
    main()
