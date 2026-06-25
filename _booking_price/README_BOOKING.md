# Booking.com Market Studio

Airbnb Market Studio 최신 개선판을 기반으로 만든 Booking.com용 시장 분석 도구입니다.

## 실행

루트 폴더에서:

```bat
Booking시장분석.bat
```

또는 이 폴더에서:

```bat
Booking시장분석.bat
```

직접 실행:

```bat
py -3 gui_app.py
```

## 모드

- A: Booking.com 검색 결과 기반 기본 Excel
- B: 기본 목록 + 상세 컬럼 Excel
- M: 평일/주말 가격 비교, 추천가, 손님용/내부용 Excel, HTML 리포트

## Booking.com 차단 대응

Booking.com은 자동 요청에 AWS/WAF 확인 페이지를 자주 반환합니다. 이 경우 프로그램이 `booking_manual_pages` 폴더에 `.url` 바로가기와 저장할 `.html` 경로를 만듭니다.

1. 생성된 `.url` 파일을 일반 Chrome 또는 Edge에서 엽니다.
2. Booking.com 확인 절차가 나오면 직접 완료합니다.
3. 실제 검색 결과 카드가 보일 때까지 기다립니다.
4. `Ctrl+S`로 안내된 `.html` 경로에 전체 페이지를 저장합니다.
5. 같은 조건으로 다시 실행합니다.

브라우저 자동 fallback을 시도하려면 Playwright 설치 후 환경변수 `BOOKING_BROWSER_FALLBACK=1`을 설정합니다.
