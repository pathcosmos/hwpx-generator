# Changelog

## [2026-04-29] Rust + rhwp 자동화 경로 추가 (cross-platform, COM 불필요)

**커밋**: `39e4070` — Add Rust+rhwp automation path (cross-platform HWP filling)
**PR**: [#1](https://github.com/pathcosmos/hwpx-generator/pull/1)

기존 **경로 A (Python + lxml + COM)** 와 나란히 동작하는 **경로 B (Rust + rhwp)** 를 추가. macOS / Linux / Windows 어디서나 한컴오피스 설치 없이 .hwp(HWP 5.0 binary) 양식을 자동 채우기 가능.

> **원칙**: from-scratch 로 새 문서를 만드는 패턴은 노출하지 않는다. 사용자 양식을 베이스로 빈 셀에만 값 삽입하는 패턴만 지원.

---

### 신규 sub-project

| 디렉토리 | 역할 |
|---------|------|
| `hwp-automate-poc/` | **Rust binary** — 기존 양식의 표 자동 채우기 데모 (헤더 매칭 → 컬럼 자동 탐색 → 라운드트립 검증) |
| `hwp-automate-py/` | **PyO3 abi3-py39 Python 바인딩** — Python 3.9~3.14 어디서나 단일 wheel 사용 |
| `hwp-automate-py/hwp_automate_cli/` | Python 보조 도구 — field_map.json 어댑터 + argparse CLI (analyze / fill / cell) |

### 외부 의존 (별도 git clone, 본 repo 비포함)

| 위치 | 출처 | 라이선스 |
|-----|------|---------|
| `../codebase/rhwp/` | [edwardkim/rhwp](https://github.com/edwardkim/rhwp) | MIT |
| `../codebase/hop/` | [golbin/hop](https://github.com/golbin/hop) (`DocumentCore::from_bytes` 패턴 출처) | MIT |

### Python API (PyO3 노출, `hwp_automate.*`)

| 함수 | 용도 |
|------|------|
| `analyze_template(path)` | 양식 표·스타일·번호 인벤토리 (read-only) |
| `fill_template(template, out, operations, dry_run=False, verify=True)` ★ | 다중 표·다중 컬럼·다중 셀 일괄 채우기. **Pre-flight + post-fill verify + dry_run** |
| `fill_template_table(template, out, mapping, ...)` | 단일 표·단일 컬럼 편의 wrapper |

### CLI

```bash
python -m hwp_automate_cli analyze --template 양식.hwp
python -m hwp_automate_cli fill    --template 양식.hwp --field-map ... --data ... --output ...
python -m hwp_automate_cli cell    --template 양식.hwp --output ... --header-match 성명 --cell 1,5,값
```

### 안전 메커니즘

- **Pre-flight 검증** — 모든 operation 의 표·컬럼·범위가 유효한지 적용 전에 확인. 잘못된 op 1개라도 있으면 양식 무수정 (silent corruption 방지).
- **Post-fill 라운드트립 검증** — 저장 후 재파싱하여 모든 셀 값이 정확히 보존됐는지 자동 확인. 불일치 시 `RuntimeError`.
- **`dry_run` 모드** — 실제 적용·저장 없이 plan 만 검증·반환.

### CLAUDE.md 업데이트

- 두 경로 비교표 (프로젝트 개요)
- 경로 B 전용 섹션 (위치, 사용법, 검증된 능력 5가지, 한계, 함정, 진입점)
- 기존 절들에 적용 범위 표기 (`HWPX 파일 수정 시 주의사항`, `COM 자동화 주의사항`)

---

### 수치 요약

| 항목 | 값 |
|------|---|
| 변경 파일 | 15 개 |
| 추가 행 | +4,189 |
| 삭제 행 | −1 |
| 신규 sub-project | 2 개 (Rust binary + Python 바인딩) |

### 검증

- 7 단계 자동 회귀 통과: analyze, fill_template 다중 op, dry_run, pre-flight 보호, CLI 호출, field_map.json 어댑터, legacy 호환
- 사용자 시각 검증 통과 (한컴에서 poc_v3.hwp 열어 확인)
- biz_plan.hwp 8 개 표 자동 발견, 5×6 인력표 자격증 4 셀 100% 라운드트립

### 알려진 한계 / 향후

- **Mac arm64 wheel 만 현재 빌드** → GitHub Actions matrix 로 macOS+Linux+Windows 자동 빌드 예정
- **rhwp v0.7.x SVG 렌더러는 outline 자동 번호 미렌더** → 한컴/모바일에서는 정상 표시
- 실 양식 검증은 사용자 양식 1 개와 함께 진행 예정

---

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
