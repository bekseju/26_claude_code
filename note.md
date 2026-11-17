+# Claude Code 학습/실습 디렉토리

이 디렉토리는 Claude Code CLI 도구를 학습하고 실습하기 위한 공간입니다. 
그리고 이 가이드는 클로드가 자동 생성한 내용입니다. 

## 디렉토리 구조 (예정)

```
26_claude_code/
├── python/          # Python 실습 프로젝트
├── typescript/      # TypeScript/JavaScript 실습 프로젝트
└── experiments/     # 자유 실험 공간
```

## 기술 스택

- **Python**: 3.11+
- **JavaScript/TypeScript**: Node.js 20+, TypeScript 5+
- **패키지 관리**: pip / npm or pnpm

## 공통 명령어

### Python

```bash
python -m venv .venv
source .venv/bin/activate      # Linux/macOS
.venv\Scripts\activate         # Windows
pip install -r requirements.txt
python main.py
```

### TypeScript

```bash
npm install
npm run dev
npm run build
npm test
```

## Claude Code 주요 기능 참고

### 슬래시 커맨드

- `/help` — 사용 가능한 커맨드 목록
- `/clear` — 대화 컨텍스트 초기화
- `/config` — 설정 변경 (모델, 테마 등)
- `/code-review` — 현재 diff 코드 리뷰
- `/init` — CLAUDE.md 자동 생성

### 유용한 패턴

- `! <command>` — 셸 명령어를 대화 컨텍스트에 직접 실행
- `/loop` — 반복 작업 자동화
- `/schedule` — 예약 실행

## 작업 방식

- 새 기능/실험은 별도 서브디렉토리에 작성
- 각 서브프로젝트에 별도 `CLAUDE.md` 추가 권장
- 실습 완료 후 핵심 내용을 주석이나 README로 정리
