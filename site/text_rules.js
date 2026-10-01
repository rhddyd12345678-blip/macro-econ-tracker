/* 자동 문장 규칙 — 카드에 들어가는 "방향/기준점" 문구와 판정 문구를
   모두 이 파일의 함수·표를 거쳐서만 만든다. 값만 고치면 사이트 전체 문구가
   함께 바뀐다.

   지켜야 할 원칙(요청 원문 그대로):
   - 좋다/나쁘다 같은 가치 판단 금지. 방향과 기준점만 말한다.
   - 인과 단정 금지: "~때문이에요" 대신 "~와 같은 시기에 나타났어요"
   - "급등·폭락·서프라이즈" 같은 과장 표현 금지.
   - 문장 속 숫자는 그래프 데이터와 같은 값에서 나와야 한다(빌드 시 검사).
   - 데이터가 없거나 오래됐으면 추측하지 말고 정해진 문구만 쓴다.
   - 예측·추천으로 읽힐 문장 금지("오를 거예요", "지금이 기회예요" 등). */

/* ---- 변화 표현 임계값(지표군별로 다르게) ---- */
var CHANGE_THRESHOLDS = {
  // 이미 %인 지표(CPI·HICP·성장률·실업률 등) — %p 기준
  default: { big: 1.0, small: 0.3 },
  // 정책금리류 — 통상 0.25%p 단위로 움직여 더 낮은 기준을 씀
  rate_level: { big: 0.5, small: 0.25 },
  // 유가·환율 등 가격류 — 상대 변화율(%) 기준
  price_level: { big: 5.0, small: 1.5 }
};

// 지표명 -> 어느 임계값 표를 쓸지
var INDICATOR_THRESHOLD_GROUP = {
  "기준금리": "rate_level",
  "정책금리": "rate_level",
  "정책금리(상단)": "rate_level",
  "WTI": "price_level",
  "두바이유": "price_level",
  "브렌트": "price_level",
  "원/달러": "price_level",
  "원/100엔": "price_level",
  "원/유로": "price_level"
};

function thresholdsFor(indicator) {
  var group = INDICATOR_THRESHOLD_GROUP[indicator] || "default";
  return CHANGE_THRESHOLDS[group] || CHANGE_THRESHOLDS.default;
}

/* ---- 방향 표현: 가치판단 없이 방향+정도만 ---- */
function describeChange(diff, indicator) {
  var t = thresholdsFor(indicator);
  if (diff >= t.big) return { text: "많이 올랐어요", dir: "up" };
  if (diff >= t.small) return { text: "올랐어요", dir: "up" };
  if (diff > -t.small) return { text: "비슷해요", dir: "flat" };
  if (diff > -t.big) return { text: "내렸어요", dir: "down" };
  return { text: "많이 내렸어요", dir: "down" };
}

/* 방향 배지에 쓰는 화살표+말(색만으로 구분하지 않기 위함) */
var DIRECTION_ARROW = { up: "▲", down: "▼", flat: "→" };

/* ---- 지표군별 비교 기준: 무엇과 비교해서 말하는지 ----
   "target": 물가안정목표(2%) 대비, "prev_change": 직전 변경 대비(정책금리),
   "prev_period": 직전 분기 대비(성장률), "prev_month": 직전 달 대비(기본값) */
var COMPARE_BASIS = {
  "CPI 헤드라인": "target", "CPI 근원": "target",
  "HICP 헤드라인": "target", "HICP 근원": "target",
  "PCE 근원": "target",
  "기준금리": "prev_change", "정책금리": "prev_change", "정책금리(상단)": "prev_change",
  "실질GDP 성장률": "prev_period"
};
var INFLATION_TARGET_PCT = 2.0;

/* ---- 금지 표현(빌드 시 카드 텍스트 전수 검사에 사용) ----
   자동 생성 문장뿐 아니라 손으로 쓴 "왜 중요할까요" 문구도 이 목록에 걸리면
   빌드가 실패하게 만든다(검사 스크립트는 구현 단계에서 추가). */
var BANNED_PHRASES = [
  "좋아요", "나빠요", "좋은", "나쁜",
  "급등", "폭락", "서프라이즈", "깜짝",
  "때문이에요", "때문입니다",
  "오를 거예요", "내릴 거예요", "오를 것", "내릴 것", "전망돼요", "예상돼요",
  "지금이 기회", "추천", "매수", "매도", "사세요", "파세요"
];

/* ---- 데이터 없음/오래됨 처리 ---- */
function fallbackText(lastDate, nextReleaseDate) {
  if (!lastDate) {
    return nextReleaseDate ? "아직 발표 전이에요(다음 발표: " + nextReleaseDate + ")" : "아직 발표 전이에요";
  }
  return "데이터를 받지 못했어요";
}

/* freq(발표 주기)에 따라 "오래된 데이터" 기준일을 판단.
   월간 45일, 분기 100일, 수시(정책금리 등)는 기준 없음(계속 최신값으로 간주). */
var STALE_DAYS_BY_FREQ = { "월간": 45, "분기": 100, "수시": null };
function isStale(lastDateStr, freq, todayStr) {
  var days = STALE_DAYS_BY_FREQ[freq];
  if (days == null) return false;
  var last = new Date(lastDateStr), today = new Date(todayStr);
  return (today - last) / 86400000 > days;
}

/* ---- 백테스트/통계 판정 문장 ----
   좋아 보이게 포장하지 않는다 — 기준 그대로 노출. */
function backtestVerdict(hitRatePct, pValue) {
  if (hitRatePct > 50 && pValue < 0.05) {
    return "이 규칙은 과거 데이터에서 어느 정도 맞았어요";
  }
  return "이 규칙은 과거 데이터에서 맞지 않았어요";
}

/* 금리 연동 판정 문장(β·R²·p값은 카드 본문에 안 쓰고 "더 알아보기"로만) */
function linkageVerdict(strongestCountryLabel, hasSignal) {
  if (!hasSignal) return "최근에는 어느 한 나라와 뚜렷하게 같이 움직인다고 보기 어려워요";
  return strongestCountryLabel + " 금리와 같은 시기에 비슷하게 움직였어요";
}

/* ---- "이번에 달라진 것" 순위 기준 ----
   지표마다 단위가 달라 그대로 비교하면 숫자가 큰 지표만 뽑힌다.
   물가·금리·성장률·실업률처럼 이미 %인 지표는 %p 절대 차이로,
   유가·환율·고용처럼 수준값인 지표는 상대 변화율(%)로 비교한다. */
var PERCENT_POINT_INDICATORS = [
  "CPI 헤드라인", "CPI 근원", "HICP 헤드라인", "HICP 근원", "PCE 근원",
  "실질GDP 성장률", "실업률", "기준금리", "정책금리", "정책금리(상단)"
];
function moverMagnitude(indicator, lastValue, prevValue) {
  if (PERCENT_POINT_INDICATORS.indexOf(indicator) !== -1) {
    return { value: lastValue - prevValue, unit: "%p" };
  }
  if (!prevValue) return null;
  return { value: (lastValue / prevValue - 1) * 100, unit: "%" };
}
