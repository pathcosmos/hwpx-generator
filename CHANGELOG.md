# Changelog

## [2026-03-08] Two-Pass Hybrid Form Filler Pipeline

**커밋**: `222eb7d` — HWPX form filler pipeline: two-pass hybrid system

기존 단순 XML 셀 채우기 + COM 찾아바꾸기 방식에서, **마크다운 문서를 파싱하여 HWPX 양식에 자동으로 채워넣는 2-Pass 하이브리드 파이프라인**으로 대폭 확장.

---

### 신규 모듈 (핵심)

| 파일 | 역할 |
|------|------|
| `src/form_filler.py` | **파이프라인 오케스트레이터** — Pass 1(XML 직접 편집) + Pass 2(COM 서식 삽입)를 순차 실행하는 메인 엔트리포인트 |
| `src/md_parser.py` | **마크다운 파서** — 사업계획서 `.md` 파일을 구조화된 블록(헤딩, 문단, 표, 리스트)으로 파싱 |
| `src/md_to_ops.py` | **마크다운→COM 변환기** — 파싱된 블록을 COM 자동화 명령(InsertText, 서식 적용 등) 시퀀스로 변환. 계층적 들여쓰기, 표→텍스트 변환 포함 |
| `src/section_mapper.py` | **섹션 매퍼** — 마크다운 섹션 번호를 HWPX 양식의 마커(##SEC1_CONTENT## 등)에 매핑 |

### 신규 템플릿

| 파일 | 역할 |
|------|------|
| `templates/gyeongnam_rbd/field_map.json` | **경남 R&BD 사업계획서 전용 필드맵** — 36개 표, 184개 셀의 좌표 매핑 (표지, 요약, 시장현황, 수요처, 목표, KPI, 연구원, 생산계획, 예산, 기관현황, 부속서류) |
| `data/form_content_map.json` | **양식 콘텐츠 매핑** — 마크다운 섹션 → HWPX 마커 대응표 |

### 신규 도구 (감사/디버깅/유틸리티)

| 파일 | 역할 |
|------|------|
| `audit_crossrefs.py` | HWPX 내 교차참조(charPrIDRef, paraPrIDRef 등) 유효성 검사 |
| `audit_hwpx_content.py` | 생성된 HWPX 파일의 콘텐츠 무결성 감사 (빈 셀, 누락 마커 탐지) |
| `audit_section0.py` | section0.xml 상세 감사 — 표/셀 구조, 텍스트 내용 점검 |
| `compare_section0.py` | 원본 vs 수정된 section0.xml 비교 (구조적 diff) |
| `compare_section0_v2.py` | 개선된 section0 비교 — 셀 단위 세밀 비교 |
| `debug_crash_isolate.py` | COM 크래시 원인 격리 — 섹션별로 나눠 실행하며 크래시 지점 특정 |
| `diagnose_xml_serialization.py` | XML 직렬화 문제 진단 — 선언부, 네임스페이스, 인코딩 검증 |
| `tools/make_rawcopy.py` | HWPX 원본 복사 유틸리티 — ZIP 엔트리별 압축방식 보존하며 클린 카피 생성 |
| `_bridge_test.py` | WSL↔Windows 브릿지 연결 테스트 |

### 기존 모듈 업데이트 (25개 파일)

#### 핵심 변경

| 파일 | 주요 변경 내용 |
|------|--------------|
| `src/bridge.py` | **COM 포스트 포맷 패턴** 구현 — `InsertText` → 선택 → `char_shape` 적용 → 해제. 0.1pt 렌더링 버그 해결. 2-tier 선택 최적화 (MoveParaBegin/End vs MoveSelLeft×N) |
| `src/hwpx_editor.py` | 다중 `hp:p`/`hp:run` 클리어링, charPrIDRef 보존 강화, 마커 삽입 기능 추가 |
| `src/hwp_com.py` | 포스트 포맷 기반 텍스트 삽입, 계층적 들여쓰기, 표→텍스트 렌더링 지원 |
| `src/extract_template.py` | 경남 R&BD 양식(38페이지, 36개 표) 분석 지원 확대 |
| `src/generate_hwpx.py` | form_filler 파이프라인 통합, 마커 기반 콘텐츠 삽입 흐름 추가 |
| `src/field_mapper.py` | 경남 R&BD 필드맵 지원, 다중 기관/기업 리스트 매핑 |

#### 분석 문서 업데이트

| 파일 | 변경 |
|------|------|
| `analysis/approach_comparison.md` | 2-Pass 하이브리드 방식 평가 결과 추가 |
| `analysis/com_evaluation.md` | 포스트 포맷 패턴 발견 및 검증 결과 기록 |
| `analysis/direct_xml_evaluation.md` | 마커 삽입 방식의 한계와 해결책 기록 |
| `analysis/hwpx_structure_analysis.md` | 경남 R&BD 양식 36개 표 구조 분석 추가 |
| `analysis/pyhwpx_evaluation.md` | 최종 평가 업데이트 |

#### 기타

| 파일 | 변경 |
|------|------|
| `.gitignore` | 출력/임시 파일 패턴 추가 |
| `CLAUDE.md` | 2-Pass 파이프라인 아키텍처, COM 포스트 포맷 주의사항 반영 |
| `README.md` | 전체 리팩토링 (아래 별도 기술) |
| `data/sample_input.json` | 경남 R&BD 양식에 맞는 샘플 데이터로 교체 |
| `data/schema.json` | 경남 R&BD 입력 스키마로 업데이트 |
| `templates/cloud_integrated/*` | 기존 템플릿 호환성 유지 보수 |
| `tests/*` | 테스트 코드 업데이트 — 새 모듈 대응 |

---

### 수치 요약

| 항목 | 값 |
|------|---|
| 변경 파일 수 | 40개 |
| 추가 행 | +11,646 |
| 삭제 행 | -5,534 |
| 순증 행 | +6,112 |
| 신규 파일 | 15개 |
| 수정 파일 | 25개 |

### Pass 1 결과 (XML 직접 편집)

- 14개 핵심 표에 184개 셀 채움
- 5개 콘텐츠 마커 삽입 (`##SEC1-5_CONTENT##`)
- 36개 표 구조 무손상, ZIP 무결성 검증 통과

### Pass 2 결과 (COM 자동화)

- 21페이지 → 65페이지로 확장 (콘텐츠 삽입)
- 5개 섹션, 총 6,396개 COM 명령 실행
- 모든 폰트 사이즈 정상 (0.1pt 버그 해결)
- 실행 시간: 약 8분
