# Project Rules & Security Guidelines (gas프로젝트)

## 🚨 [최우선 원칙] API 키 및 시크릿 보안 관리 (Zero-Leak Policy)

이 프로젝트(Google Apps Script 기반 WMS)를 진행할 때 API 키 및 민감 정보의 보안을 최우선 조건으로 준수합니다.

1. **소스코드 및 HTML 내 하드코딩 절대 금지 (No Hardcoding):**
   - `Code.js`, `*.html` 등 어떤 소스코드에도 실제 Gemini API Key나 토큰을 평문 또는 Base64 인코딩 형태로 하드코딩하지 않습니다.
   - Base64 인코딩은 누구나 디코딩할 수 있으므로 Fallback 용도로도 절대 코드에 삽입하지 않습니다.

2. **클라우드 환경변수(스크립트 속성) 활용:**
   - Google Apps Script 환경에서는 모든 민감 API 키를 `PropertiesService.getScriptProperties()`(스크립트 속성)에 등록하여 안전하게 불러옵니다.
   - 로컬 보관용 `.env` 파일은 절대 Git에 커밋되지 않도록 `.gitignore`에 등록하여 관리합니다.

3. **Git 커밋 및 원격 푸시 전 검증:**
   - 커밋 전에 `git diff` 및 `git status`를 확인하여 시크릿 정보나 설정 파일(`.env`, `.clasprc.json` 등)이 포함되지 않았는지 철저히 검증합니다.

4. **저장소 접근 권한 관리:**
   - 회사 내부 WMS 로직이 포함된 저장소는 GitHub 설정을 Private으로 유지하여 외부 노출을 방지합니다.
