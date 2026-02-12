# bid_g2b_data
## 1.api 클라이언트 패턴
```
class G2BAPIClient:
  def __init__(self, service_key):
    self.base_url="..."
    self.service_key=service_key
```
1) api 호출 로직을 클래스로 캡슐화
2) 재사용 가능하고 유지보수가 쉬운 구조
3) 나라장터 api의 인증키와 기본 url을 객체로 관리

## 2. 페이징 처리
```
def fetch_all_pages(self,search_params):
  while True:
    data=self.fetch_bid_notices(search_params)
    if not data or len(data) < num_of_rows:
        break
    page_no +=1
```
1) api는 한번에 최대 100건만 반환
2) 자동으로 다음 페이지를 계속 호출해서 전체 데이터 수집
3) time.sleep(0.1)로 api서버 부하 방지

## 3. 날짜 범위 자동 분할
```
def get_automatic_data_ranges():
  start_date=(now-timedelta(weeks=2))
  while current_end=current_start+timedelta(days=29)
```
1) 나라장터 api는 한번에 조회할 수 있는 기간 제한이 있을 수 있음
2) 2주 기간을 30일 단위로 잘게 쪼개서 안전하게 조회
3) 월이 바뀌는 경우도 대응
