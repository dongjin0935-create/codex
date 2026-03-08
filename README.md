# 강남조선 발주서 자동 변환 프로그램

## 설치 방법
```bash
python -m venv .venv
. .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 실행 방법
```bash
python main.py \
  --pdf "C:\\Users\\김재현\\Desktop\\gangnam_project\\20260129-224001-0154-동진산업.pdf" \
  --erp-template "C:\\Users\\김재현\\Desktop\\gangnam_project\\224001-0154.xlsx" \
  --trade-template "C:\\Users\\김재현\\Desktop\\gangnam_project\\표준거래명세표(R.2)_동진산업.xls" \
  --output-dir "C:\\Users\\김재현\\Desktop\\gangnam_project\\output" \
  --config "C:\\Users\\김재현\\Desktop\\gangnam_project\\config\\generated_mapping.json" \
  --trade-ext xls
```

## 분석 전용 실행
```bash
python main.py --pdf "...pdf" --erp-template "...xlsx" --trade-template "...xls" --output-dir "...\\output" --analyze-only
```

## 필수 패키지
- pdfplumber (PDF 텍스트 추출)
- PyMuPDF (추가 fallback 용)
- openpyxl (ERP xlsx 작성)
- pywin32 (Windows Excel COM 자동화, xls 템플릿 보존 핵심)
- pandas, pyyaml
- pytest

## Windows에서 Excel 설치가 필요한 이유
`.xls` 템플릿의 병합셀/서식/수식을 최대한 보존하려면 Excel COM 자동화가 가장 안정적입니다. 이 프로젝트는 pywin32 + Excel.Application 방식이 기본입니다.

## xls 템플릿 처리 권장 방식
1. 템플릿 xls를 COM으로 열기
2. SaveAs로 출력 파일 생성
3. 메타/품목 데이터 입력
4. 저장

## 출력 파일 예시
- `강남거래명세서_20260129_224001_0154.xls`
- `ERP업로드_20260129_224001_0154.xlsx`

## 디버깅 방법
- `--debug` 옵션 사용
- 로그 파일: `output/logs/run.log`

## 프로젝트 구조
```text
gangnam_project/
  main.py
  requirements.txt
  README.md
  config/
    constants.py
    generated_mapping.json
  src/
    models.py
    parser/
      filename_parser.py
      pdf_parser.py
    analyzer/
      erp_analyzer.py
      trade_template_analyzer.py
      mapping_builder.py
    writers/
      erp_writer.py
      trade_writer_base.py
      trade_writer_win32.py
      trade_writer_openpyxl.py
    utils/
      date_utils.py
      number_utils.py
      excel_utils.py
      logging_utils.py
      file_utils.py
  tests/
    test_filename_parser.py
    test_pdf_parser.py
    test_mapping_builder.py
    test_utils.py
```
