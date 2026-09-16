# macro-regime-watch — 변곡점 대시보드

재무부·연준 개입 체제의 변곡 신호를 신호등으로 판독하는 개인 리서치 보조 대시보드.

**대시보드:** https://hoee213.github.io/macro-regime-watch/

## 감시 축

| 축 | 지표 | 판독 |
|---|---|---|
| 개입 체제 | FIMA 레포, 중앙은행 스왑, 엔/달러 | FIMA 가동 + 스왑 동반 = 스트레스성(B), FIMA 단독 = 개입성(A) |
| 장기금리 박스권 | 10Y·30Y, 장기물 입찰 | 30Y 5.2% + 입찰 부진 = 베센트 반응함수 존 |
| 인플레·재정우위 | 5y5y BE, ACM 텀프리미엄 | 5y5y 2.6% 또는 TP 급등 = 재정우위 프라이싱 |
| AI 크레딧 | AA OAS vs IG OAS | AA 상대 확대 = 하이퍼스케일러 크레딧 프록시 |

## 소스

- **재무부 (home.treasury.gov)** — 10Y·30Y 파 수익률, 5y5y 브레이크이븐(명목·실질 곡선에서 FRED T5YIFR 공식으로 계산)
- **연준 H.4.1 데이터패키지** — FIMA 레포, 중앙은행 스왑, 총자산, 외국공적 역레포 풀
- **뉴욕연준** — ACM 10년 텀프리미엄 (xls, `xlrd` 필요)
- **FRED** — ICE BofA OAS(IG·AA) 전용. GitHub 러너 IP는 fred.stlouisfed.org가 차단하므로
  `FRED_API_KEY` 시크릿이 있을 때만 수집되고, 없으면 해당 카드만 결측으로 표시된다.
  키는 https://fredaccount.stlouisfed.org/apikeys 에서 무료 발급.
- **TreasuryDirect** — 10Y/20Y/30Y 낙찰 결과 (재발행 포함)
- **Yahoo Finance** — 엔/달러 실시간 (FRED DEXJPUS는 4영업일 지연이라 폴백용)

## 실행

```bash
pip install requests xlrd
python macro_dashboard.py                 # output/index.html
python macro_dashboard.py --out docs --json
python macro_dashboard.py --telegram --url <URL>   # 요약 발송
```

## 자동화

`.github/workflows/dashboard.yml` — 매일 07:00 KST (22:00 UTC) 실행 → `docs/` 갱신 → Pages 배포.

텔레그램 요약은 `TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID` 시크릿이 등록돼 있을 때만 발송된다.
링크를 붙이려면 저장소 변수 `PAGE_URL`을 설정한다.

수치 확정치는 원 소스 기준이며 본 대시보드는 개인 리서치 보조용이다.
