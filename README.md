## AWS, RDS 리소스 정리 자동화

### 컴플라이언스 대비 자산관리 구글시트
- 구글 api는 1분에 100번까지 호출가능해서 한라인씩 입력하여 콜을 최소화
- `time.sleep()`으로 rate limit에 걸리지 않도록 함