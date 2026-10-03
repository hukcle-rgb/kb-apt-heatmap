# 익산시 망성면 무형리 27필지 평가액 조회

`vworld_valuation.py`는 브이월드 오픈API에서 필지마다 지목, 면적(㎡), 2026년 개별공시지가(원/㎡)를 조회합니다.
그 결과로 `무형리_27필지_평가액.xlsx`를 만듭니다. 평가액은 `면적 × 공시지가` 수식으로 계산하고, 맨 아래에 합계 행이 들어갑니다.

```bash
pip install openpyxl requests
python vworld_valuation.py --key <브이월드_인증키>              # 조회 후 엑셀 생성
python vworld_valuation.py --key <인증키> --domain <등록URL>     # 키에 서비스 URL이 묶여 있는 경우
python vworld_valuation.py --offline                             # 빈 양식만 생성
```

- PNU는 `5214038026`(법정동코드) + `1`(일반) + 본번 4자리 + 부번 4자리로 만듭니다.
  - 무형리의 법정동코드는 원래 `4514038026`입니다. 2024-01-18 전북특별자치도가 출범하면서 시도코드가 45에서 52로 바뀌었습니다.
- 조회 결과가 없으면 산 필지(`2`)로 다시 조회합니다. 그래도 없으면 예전 코드(`45…`)로 조회합니다. 이렇게 대체 조회한 필지는 비고 칸에 표시됩니다.
- 사용하는 API
  - 지목·면적: `ned/data/ladfrlList`(토지임야목록)
  - 공시지가: `ned/data/getIndvdLandPriceAttr`(개별공시지가, stdrYear=2026)
