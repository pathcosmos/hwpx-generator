# hwp-automate-poc — Rust PoC

`hwpx-generator` 의 **경로 B (Rust + rhwp)** 에 속하는 검증용 binary 모음. 한컴오피스 설치 없이 macOS/Linux/Windows 어디서나 .hwp(HWP 5.0 binary) 를 자동 생성·편집할 수 있음을 코드 + dump + 시각(한컴) 검증으로 입증한 PoC.

상위 컨텍스트는 `../CLAUDE.md` 의 [Rust + rhwp 경로](../CLAUDE.md#rust--rhwp-경로-크로스플랫폼-com-불필요) 섹션 참조.

## 무엇을 검증했나

다섯 가지 능력이 코드로 가능함을 증명:

1. **표 생성 + 헤더 행 표시** — `repeat_header=true` 자동, 12셀 round-trip 일치
2. **스타일 복사·적용** — 헤더 셀과 데이터 셀의 ParaShape 분리 (apply_cell_style_native)
3. **항목 자동 번호** — `head=Outline level=0` IR (한컴 표준 패턴)
4. **기존 문서 같은 스타일** — `DocumentCore::from_bytes` 로 22~26개 스타일 그대로 사용
5. **기존 양식 표에 값 정확 삽입** — 헤더 매칭으로 컬럼 자동 탐색, biz_plan.hwp 5×6 인력표 자격증 4셀 100% 일치

### V1 — 실 양식 검증 통과 (PR #2)

| 양식 | 크기 | 결과 |
|------|------|------|
| YCP_V0.4 (제조AI 사업신청서) | 35MB → 35MB | ✅ applied + verified, 6 셀 채움 |
| 코리녹스_V0.9 (제조AI 사업신청서) | 54MB → 54MB | ✅ applied + verified, 5 셀 채움 |

`preserve_images=True` (기본) 로 **BinData 54/54 동일 크기 보존** — 한컴이 손상으로 인식하지 않음. 사용자 시각 검증 통과 ("디게 괜찮다").

PoC 단계가 끝나고 PoC v3 의 패턴 (양식 자동 채우기 binary) 이 본 main.rs 로 승격됨. from-scratch 빌드 패턴은 의도적으로 노출하지 않음 — 사용자 양식만 채운다는 원칙.

## 디렉토리

```
hwp-automate-poc/
├── Cargo.toml             rhwp = ../../codebase/rhwp 의존
├── src/
│   ├── main.rs            PoC v2 binary (새 문서 처음부터)
│   └── bin/
│       └── fill_template.rs  PoC v3 binary (양식 표 채우기)
├── output/                생성된 .hwp / .svg (gitignore 됨)
└── target/                cargo build 산출물 (gitignore 됨)
```

## 빌드 & 실행

### 사전 요구사항

- Rust 1.75+ (`brew install rust` 또는 rustup)
- `../../codebase/rhwp/` 위치에 [edwardkim/rhwp](https://github.com/edwardkim/rhwp) 가 git clone 으로 존재

```bash
# 만약 아직 codebase 가 없다면
mkdir -p ../../codebase && cd ../../codebase
gh repo clone edwardkim/rhwp
cd -
```

### 빌드

```bash
cargo build              # dev 빌드 (~20초 첫 빌드, 이후 ~1초)
cargo build --release    # release 빌드 (LTO, 더 느림)
```

### 실행 — `cargo run`

```bash
cargo run                                                    # 기본: biz_plan.hwp 의 자격증 컬럼 자동 채우기
cargo run -- ../../codebase/rhwp/samples/exam_kor.hwp        # 다른 템플릿 지정
cargo run -- <template.hwp> <output.hwp>                     # 입출력 모두 지정
```

수행 흐름:
1. 사용자 양식 로드 (`DocumentCore::from_bytes`)
2. Document IR 순회로 모든 표 자동 발견
3. 헤더 행에 "성명" 텍스트가 포함된 5×6 표 자동 식별 (인덱스 하드코딩 없음)
4. "자격증" 컬럼 인덱스 자동 탐색
5. 4명 각자 자격증 셀에 값 삽입 (`insert_text_in_cell_native`)
6. `output/poc_v3.hwp` 저장 (~28 KB)
7. 라운드트립 검증 — 4/4 셀의 기대값/실제값 일치

실무 자동화의 핵심 패턴: **양식 로드 → 헤더 매칭 → 컬럼 자동 탐색 → row 단위 채움**. 양식이 살짝 변해도 (컬럼 순서, 행 추가) 코드 변경 없이 견딤.

> **중요 원칙**: 새 문서를 from-scratch 로 만드는 패턴은 본 PoC 에서 사용하지 않는다. 항상 사용자가 제공한 양식을 베이스로 한다. (`create_blank_document_native` / blank 템플릿 직접 빌드는 학습용으로도 권장하지 않음.)

## 결과물 검증

### 한컴/모바일 한글에서 시각 확인 (가장 강한 검증)

`output/poc_v3.hwp` 를 한컴 한글, 한컴 모바일, 또는 한컴 웹기안기에 업로드해서 열어보면:
- 8개 표 모두 정상
- 인력 명단 표 (이O수/박O호/조O현/전O호) 자격증 컬럼에 4개 값
- 다른 셀·텍스트·스타일 무손상

### rhwp CLI 로 IR 검증

```bash
# rhwp CLI 빌드 (1회)
cd ../../codebase/rhwp && cargo build --bin rhwp && cd -

RHWP=../../codebase/rhwp/target/debug/rhwp
$RHWP info output/poc_v3.hwp           # 파일 메타·표 통계
$RHWP dump output/poc_v3.hwp           # 전체 IR (스타일 ID, ParaShape, Cell 텍스트)
$RHWP export-svg output/poc_v3.hwp -o output/svg/   # SVG 렌더링
```

`rhwp dump` 출력에서 우리 입력값이 정확히 나타나면 round-trip 통과.

## 핵심 코드 패턴

### 1. 사용자 양식 로드 (시작점, 항상)

```rust
let bytes = fs::read("template.hwp")?;
let mut core = DocumentCore::from_bytes(&bytes)?;  // hop 패턴 흡수
```

### 2. 셀 채우기 (배치 모드)

```rust
core.begin_batch_native()?;
for (cell_idx, text) in ... {
    core.insert_text_in_cell_native(sec, para_idx, ctrl_idx, cell_idx, 0, 0, text)?;
}
core.end_batch_native()?;
```

### 3. 스타일 적용 (스타일 테이블 인덱스)

```rust
let outline1_id = core.document().doc_info.styles.iter()
    .position(|s| s.local_name == "개요 1")
    .ok_or("스타일 없음")?;
core.apply_cell_style_native(0, para_idx, 0, cell_idx, 0, outline1_id)?;
core.apply_style_native(0, body_para_idx, outline1_id)?;
```

### 4. 표 자동 발견 (양식 분석)

```rust
for (sec_idx, section) in core.document().sections.iter().enumerate() {
    for (para_idx, para) in section.paragraphs.iter().enumerate() {
        for (ctrl_idx, ctrl) in para.controls.iter().enumerate() {
            if let Control::Table(t) = ctrl {
                // t.row_count, t.col_count, t.cells[i].row/col/paragraphs
            }
        }
    }
}
```

### 5. 저장

```rust
let bytes = core.export_hwp_native()?;
fs::write("output.hwp", &bytes)?;
```

## 한계 (현 시점, rhwp v0.7.x)

- **outline 자동 번호 SVG 미렌더** — IR 은 정확하나 rhwp 의 SVG 렌더러는 `head=Outline` 의 자동 번호 텍스트를 안 그림. 한컴/모바일에서 열면 정상.
- **HWPX 직렬화 표/그림 부분 미완** — 출력은 HWP 5.0 binary 권장.
- **DocumentCore.document 가 `pub(crate)`** — 외부에서 IR 직접 변형 불가. 모든 변경은 `*_native` 메서드 경유.
- **`apply_style_native` 가 새 ParaShape 등록** — 호출마다 풀이 약간씩 늘 수 있음 (실무 영향은 미미).

## 다음 확장 후보

- 다양한 양식에 대한 헤더 매칭 robustness 테스트
- `field_map.json` (기존 경로 A 형식) 호환 어댑터
- CLI 옵션 확장 (`--template`, `--data <json>`, `--output`)
- Windows / Linux 에서의 회귀 검증

## Acknowledgement (감사·출처)

본 PoC 는 두 개의 외부 오픈소스 프로젝트 위에 만들어졌습니다. 우리는 그 위에 얇은 자동화 레이어를 더했을 뿐이며, HWP 처리의 핵심 가치는 모두 다음 프로젝트들의 결과물입니다.

### 🦀 rhwp — 핵심 엔진 (직접 의존)

- **저자:** Edward Kim ([@edwardkim](https://github.com/edwardkim))
- **저장소:** https://github.com/edwardkim/rhwp
- **라이선스:** MIT
- **설명:** Rust + WebAssembly 기반 오픈소스 HWP/HWPX 뷰어/에디터. v0.7.x 시점 891+ 테스트, hyper-waterfall 방법론으로 개발 (작업지시자-AI 페어 프로그래밍).

**본 PoC 가 사용하는 rhwp 의 모듈·기능:**

| rhwp 모듈/기능 | 본 PoC 가 사용하는 방법 |
|---|---|
| `rhwp::document_core::DocumentCore` | 양식 로드(`from_bytes`), 셀 텍스트 삽입, 표 생성, 스타일 적용 등 IR 조작의 핵심 API |
| `rhwp::parser::parse_document` | HWP 5.0 / HWPX 자동 포맷 감지 + 파싱 |
| `rhwp::serializer::serialize_hwp` | Document IR → HWP 5.0 binary 출력 (`export_hwp_native` 가 호출) |
| `rhwp::parser::cfb_reader::LenientCfbReader` | 비표준 CFB 메타데이터를 가진 양식 파일도 lenient 파싱 |
| `rhwp::serializer::mini_cfb::build_cfb` | rhwp 의 mini CFB writer — 한컴 호환 CFB v3 컨테이너 작성 |
| `rhwp::model::control::Control::Table`, `model::table::{Table, Cell}` | 표/셀 모델 — `discover_tables` 에서 IR 직접 순회 |

**의존 방식:** 본 PoC 의 `Cargo.toml` 에 `rhwp = { path = "../../codebase/rhwp", default-features = false }` 로 path 의존. **rhwp 코드는 일절 수정하지 않음** — upstream 그대로 사용. rhwp 업데이트 시 같은 위치에 새 버전 clone 으로 교체 가능.

**왜 rhwp 인가:** Mac/Linux/Windows 어디서든 한컴오피스 설치 없이 .hwp 처리가 가능한 유일한 오픈소스 엔진. 891+ 테스트가 IR 라운드트립 무결성을 보장하고, 한컴이 받아들이는 valid HWP 5.0 바이너리를 생성합니다. 본 PoC 의 모든 신뢰성은 rhwp 의 IR 모델·파싱·직렬화 정확도에 기반합니다.

### 🪝 hop — 패턴 출처 (참고만, 직접 의존 안 함)

- **저자:** golbin ([@golbin](https://github.com/golbin))
- **저장소:** https://github.com/golbin/hop
- **라이선스:** MIT
- **설명:** Tauri 2 기반 macOS/Windows/Linux 데스크톱 HWP 뷰어·에디터. rhwp 를 엔진(third_party/rhwp 서브모듈)으로 사용.

**본 PoC 가 hop 에서 흡수한 패턴:**

| hop 의 코드 | 본 PoC 가 영향받은 부분 |
|---|---|
| `apps/desktop/src-tauri/src/state.rs` 의 `editable_core_from_bytes` | 양식을 `DocumentCore::from_bytes()` 한 줄로 로드하는 표준 진입점 패턴 |
| `commands.rs` 의 `mutate_document(operation, args)` JSON dispatcher | 우리 `fill_template(operations=[...])` 의 다중 op 디스패치 설계 영감 |
| `commit_staged_hwp_save` (atomic save: staged → rename) | 향후 운영 시 사용자 양식 보호용 staged write 패턴 (현재는 우리 출력은 별도 경로라 생략) |
| 의도적으로 `rhwp::DocumentCore` 만 사용, IR 의 내부 필드는 안 만짐 | rhwp 와의 깔끔한 boundary 유지 — 우리도 동일 정책 |

**의존 방식:** **직접 코드 의존 없음**. 단지 hop 의 코드를 읽고 좋은 패턴을 본 PoC 설계에 흡수. hop 자체는 GUI 앱이라 우리 헤드리스 자동화 흐름과는 다른 영역.

**왜 hop 패턴인가:** rhwp 를 운영 환경에서 어떻게 쓰는지 보여주는 가장 완성도 높은 오픈소스 사례. 직접 코드를 읽으며 "raw IR 만지지 말고 `*_native` 메서드 경유" 같은 안전 정책을 자연스럽게 흡수했습니다.

### ⚙️ 그 외 사용 라이브러리

| 라이브러리 | 용도 | 라이선스 |
|---|---|---|
| Rust toolchain (`cargo`, `rustc`) | 빌드 | MIT/Apache-2.0 |
| rhwp 의 transitive 의존: cfb, byteorder, zip, quick-xml, encoding_rs, image, usvg, ttf-parser 등 | rhwp 가 직접 사용 (우리는 transitive) | 각 crate 의 MIT/Apache-2.0 |

### 한글 / 한컴 상표 안내

- **"한글", "한컴", "HWP", "HWPX"** 는 주식회사 한글과컴퓨터의 등록 상표입니다.
- 본 프로젝트(hwpx-generator 의 hwp-automate-poc 서브프로젝트)는 한글과컴퓨터와 제휴, 후원, 승인 관계가 없는 **독립적인 오픈소스 작업**입니다.
- HWP 5.0 바이너리 포맷 처리는 rhwp 가 한글과컴퓨터의 공개 문서를 참고하여 구현한 결과를 활용합니다.

### 본 PoC 의 위치·범위

- **목적:** rhwp 엔진을 사용한 양식 자동 채우기의 가능성 검증 + Python 바인딩 전제 구현
- **상태:** PoC (proof-of-concept). 사용자 내부 업무 자동화 목적
- **외부 배포·재배포 시:** rhwp / hop / Hangul/Hancom 상표 표기 의무 준수 필수
