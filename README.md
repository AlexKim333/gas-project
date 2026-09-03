# gas-project (GAS 기반 WMS 프로젝트)

Google Apps Script(GAS) 및 구글 스프레드시트를 기반으로 구축된 창고 재고 관리 시스템(WMS)입니다.

## 📁 파일 구성

- **Code.js**: GAS 백엔드 서버 로직 (재고 조회, 제품 등록, 입출고 처리, 데이터 동기화 등)
- **WarehouseApp.html**: WMS 메인 웹 앱 UI (입출고, 재고 관리 등)
- **SearchModify.html**: 데이터 검색 및 수정 UI
- **searchform.html**: 검색 폼 인터페이스
- **ppsscript.json**: Google Apps Script 프로젝트 매니페스트 설정
- **.clasp.json**: Clasp 연동 설정 (Script ID 포함)

## 📊 연동 구글 시트 구조

- 재고시트: 품목, 색상, 박스재고, 낱개재고, 안전재고, 박스당수량, 초기재고, 메이커 등
- PendingSheet: 보류/처리 대기 내역
- 입고처목록: 입고처 정보
- 출고처목록: 출고처 정보
- 관리자명단: 승인된 관리자 목록
- 메이커: 제조사/브랜드 목록
