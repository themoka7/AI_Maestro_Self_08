# 과제 제출 트래커 (6회차 · GAS 자동화 앱)

Google Apps Script 웹 앱 하나로 **과제 제출 접수 → 시트 기록 → 회차별 집계 → 매일 알림 메일**까지 자동화한 앱입니다.
화면은 GitHub Pages 정적 페이지이고, 백엔드는 GAS 웹 앱 + Google Sheets 입니다.

- 페이지: https://themoka7.github.io/AI_Maestro_Self_08/

웹 앱 URL 을 연결하기 전에는 페이지가 **데모 모드**(브라우저 localStorage)로 동작합니다.

## 구조

```
gas/
  Code.gs          # doGet/doPost API, setup, 트리거, 일일 알림
  appsscript.json  # 타임존 Asia/Seoul, 웹 앱(배포자 권한·익명 접근)
site/
  index.html       # 제출 폼 + 현황 대시보드 (의존성 없음)
  config.js        # GAS 웹 앱 URL
.github/workflows/
  pages.yml        # site/ → GitHub Pages
  gas-deploy.yml   # gas/ → Apps Script (clasp, 시크릿 있을 때만)
```

## 동작

| 단계 | 구현 |
|---|---|
| 제출 | 페이지 → `doPost` (`Content-Type: text/plain` 으로 CORS preflight 회피) |
| 기록 | `Submissions` 시트에 append, `Roster` 에 이메일 기준 자동 등록. `LockService` 로 동시 제출 직렬화 |
| 집계 | `doGet?action=summary&round=N` — 회차별 사람당 최신 제출 기준. 공개 응답의 이메일은 마스킹 |
| 알림 | `sendDailyDigest` 매일 09시: 운영자에게 요약, 마감 3일 전부터 미제출자에게 리마인더 |

API

```
GET  <exec>?action=summary&round=6
GET  <exec>?action=recent&limit=20
GET  <exec>?action=ping
POST <exec>   {"action":"submit","round":6,"name":"","email":"","title":"","link":"https://...","note":""}
```

## GAS 배포 (최초 1회)

1. https://script.google.com 에서 새 프로젝트 생성 → `gas/Code.gs`, `gas/appsscript.json` 내용 붙여넣기
   (편집기 설정에서 "appsscript.json 표시" 체크)
2. 함수 `setup` 실행 → 권한 승인 → 로그에 DB 스프레드시트 URL 출력
3. 함수 `installTriggers` 실행 → 매일 09시 알림 트리거 등록
4. 프로젝트 설정 → 스크립트 속성
   - `CURRENT_ROUND` : `6` (setup 이 기본값 설정)
   - `DEADLINE` : 예) `2026-10-19T23:59:00+09:00` — 없으면 학생 리마인더 미발송
   - `OWNER_EMAIL` : 요약 메일 수신자 (기본: 배포자)
   - `REMIND_STUDENTS` : `false` 면 학생 리마인더 끔
5. 배포 → 새 배포 → 유형 **웹 앱**, 실행: 나, 액세스: 모든 사용자 → `/exec` URL 복사
6. `site/config.js` 의 `GAS_API_URL` 에 URL 입력 후 push → Pages 재배포
   (커밋 없이 확인하려면 `?api=<exec URL>` 쿼리)

### CI 로 GAS 자동 배포 (선택)

저장소 시크릿에 `CLASPRC_JSON`(`npx @google/clasp login` 후 `~/.clasprc.json`), `GAS_SCRIPT_ID` 를 넣으면
`gas/` 변경 push 시 `clasp push` + `deploy` 를 수행합니다. `GAS_DEPLOYMENT_ID` 를 넣으면 같은 `/exec` URL 로 버전만 갱신합니다.
사용자 계정에서 [Apps Script API](https://script.google.com/home/usersettings) 를 켜 두어야 합니다.

## Pages 설정

Settings → Pages → Source 를 **GitHub Actions** 로 지정하면 `main` 에 `site/` 변경을 push 할 때 자동 배포됩니다.
(`github-pages` 환경은 기본적으로 `main` 만 배포를 허용합니다.)
