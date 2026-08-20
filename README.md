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

- **FRED** (API 키 불필요) — H.4.1 주간 계열, 일간 금리·스프레드
- **TreasuryDirect** — 10Y/20Y/30Y 낙찰 결과 (재발행 포함)
- **Yahoo Finance** — 엔/달러 실시간 (FRED DEXJPUS는 4영업일 지연이라 폴백용)

## 실행

```bash
pip install requests
python macro_dashboard.py                 # output/index.html
python macro_dashboard.py --out docs --json
python macro_dashboard.py --telegram --url <URL>   # 요약 발송
```

## 자동화

`.github/workflows/dashboard.yml` — 매일 07:00 KST (22:00 UTC) 실행 → `docs/` 갱신 → Pages 배포.

텔레그램 요약은 `TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID` 시크릿이 등록돼 있을 때만 발송된다.
링크를 붙이려면 저장소 변수 `PAGE_URL`을 설정한다.

수치 확정치는 원 소스 기준이며 본 대시보드는 개인 리서치 보조용이다.
