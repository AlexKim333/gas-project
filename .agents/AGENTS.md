# Antigravity Agent Rules - gas프로젝트

1. **API 키 보안 최우선 (Zero-Leak Security):**
   - 어떠한 경우에도 소스코드(`Code.js`, `*.html` 등) 내에 실제 API Key, Secret, Token을 하드코딩(평문 및 Base64 인코딩 포함)하지 않습니다.
   - API 키는 오직 Google Apps Script의 스크립트 속성(`PropertiesService.getScriptProperties()`)을 통해서만 취득합니다.

2. **환경변수 파일 Git 유출 방지:**
   - 로컬 `.env`, `.env.*`, `.clasprc.json` 파일이 `.gitignore`에 등록되어 있는지 항상 검증하고, Git 커밋에 포함되지 않도록 차단합니다.

3. **기존 기능 파괴 방지 (No Regression):**
   - 기존의 재고 관리, 입출고, 검색수정, 바코드, OCR 기능이 패치로 인해 손상되지 않도록 영향도를 사전 점검합니다.
