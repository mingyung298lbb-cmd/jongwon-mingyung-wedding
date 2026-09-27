# 웨딩 RSVP / 댓글 Google Sheets 연결

## 시트
아래 두 탭을 사용합니다.

### RSVP
3행 헤더:
제출시간 | 구분 | 성함 | 연락처 | 참석여부 | 참석인원 | 고유키 | 메모

### 댓글
3행 헤더:
제출시간 | 성함 | 댓글 | 공개여부 | 관리메모

- 새 댓글은 자동으로 `숨김`으로 저장됩니다.
- 청첩장에는 공개여부를 `공개`로 바꾼 댓글만 표시됩니다.
- RSVP 데이터는 공개 조회 API가 없습니다.

## Apps Script 배포
1. Google Sheet에서 **확장 프로그램 → Apps Script**
2. `rsvp-apps-script.gs` 내용을 붙여넣기
3. 코드의 `SPREADSHEET_ID`를 현재 Google Sheet ID로 변경
4. 프로젝트 설정에서 시간대를 **Asia/Seoul**로 설정
5. **배포 → 새 배포 → 웹 앱**
6. 실행 사용자: **나**
7. 액세스 권한: **모든 사용자**
8. 배포 후 생성된 `https://script.google.com/macros/s/.../exec` 주소를 복사
9. `index.html`의 `RSVP_API_URL`에 붙여넣기

Google Sheet 자체 공유 설정은 **제한됨** 상태로 유지합니다.
웹 앱은 제출을 받아 시트에 쓰지만, RSVP 목록은 외부에 반환하지 않습니다.
