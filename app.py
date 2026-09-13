import json
import re
from pathlib import Path

import pandas as pd
import streamlit as st
from openai import OpenAI

st.set_page_config(
    page_title="수강후기 AI 분석 리포트",
    page_icon="📊",
    layout="wide",
)

APP_TITLE = "수강후기 AI 분석 리포트"
MODEL = "gpt-5.6-luna"
SAMPLE_FILE = Path(__file__).with_name("reviews_seoul_center_final.csv")
MAX_REVIEWS_PER_COURSE = 150

st.title(f"📊 {APP_TITLE}")
st.caption(
    "수강후기 데이터를 업로드하면 과정별 핵심 키워드, 감성분석, 주요 의견, "
    "개선방안 및 개인별 피드백을 생성합니다."
)

try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error(
        "OpenAI API 키가 설정되지 않았습니다. "
        "Streamlit Cloud의 App settings → Secrets에 "
        'OPENAI_API_KEY = "sk-..." 형식으로 등록해주세요.'
    )
    st.stop()


def normalize_name(x):
    return str(x).strip().lower().replace(" ", "").replace("_", "")


def find_col(df, candidates):
    normalized = {col: normalize_name(col) for col in df.columns}
    for candidate in candidates:
        c = normalize_name(candidate)
        for col, ncol in normalized.items():
            if c == ncol or c in ncol or ncol in c:
                return col
    return None


def load_dataframe(file_obj):
    name = file_obj.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(file_obj)
        except UnicodeDecodeError:
            file_obj.seek(0)
            return pd.read_csv(file_obj, encoding="cp949")
    return pd.read_excel(file_obj)


@st.cache_data(show_spinner=False)
def load_sample(path_str):
    return pd.read_csv(path_str)


def clean_dataframe(df, course_col, review_col, rating_col):
    work = df.copy()
    work = work.dropna(subset=[course_col, review_col])
    work[course_col] = work[course_col].astype(str).str.strip()
    work[review_col] = work[review_col].astype(str).str.strip()
    work = work[work[review_col].str.len() >= 3]
    if rating_col:
        work[rating_col] = pd.to_numeric(work[rating_col], errors="coerce")
    return work.reset_index(drop=True)


def extract_json(text):
    if not text:
        raise ValueError("모델 응답이 비어 있습니다.")

    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("JSON 객체를 찾지 못했습니다.")
    return json.loads(text[start:end + 1])


def safe_list(obj, key, limit=None):
    value = obj.get(key, [])
    if not isinstance(value, list):
        return []
    return value[:limit] if limit else value


PROMPT_TEMPLATE = """
당신은 직업훈련기관의 강의평가·수강후기를 분석하는 데이터 분석가입니다.

아래는 하나의 교육과정에 대한 수강생 후기입니다.
각 후기에는 [리뷰번호], 선택적으로 [별점], [후기본문]이 있습니다.

[과정명]
{course_title}

[분석대상 후기 수]
{n_reviews}

[후기 목록]
{reviews_block}

아래 기준을 반드시 지켜 분석하십시오.

1. 핵심 키워드 정확히 10개
- 단순 조사·일반어가 아니라 교육과정 개선에 의미가 있는 명사/명사구 중심
- frequency는 해당 키워드 또는 같은 의미의 표현이 등장한 '후기 건수'를 기준으로 가능한 한 정확하게 계산
- meaning은 이 교육과정에서 그 키워드가 어떤 의미인지 한 문장으로 설명

2. 감성 분석
- 모든 분석대상 후기를 긍정/중립/부정 중 하나로 분류했다고 가정하고 합계가 반드시 {n_reviews}가 되게 함
- 긍정, 중립, 부정의 주요 사유를 각각 최대 5개 제시
- 해당 감성이 거의 없다면 억지로 지어내지 말고 빈 배열 허용

3. 개선방안
- 후기 근거를 바탕으로 구체적이고 실행 가능한 개선방안 정확히 5개

4. 개인별 피드백
- 대표성이 높은 후기 최대 10개를 골라 작성
- review_no에는 아래 후기 목록의 실제 리뷰번호를 사용
- feedback은 해당 수강생에게 전달 가능한 정중하고 구체적인 메시지
- 단순 칭찬 반복보다 의견 반영/추가 학습/보완 방향이 드러나게 작성

중요:
- 후기에서 확인되지 않는 사실을 만들지 마십시오.
- 개인정보를 추정하거나 생성하지 마십시오.
- 반드시 아래 JSON 구조만 반환하십시오.
- 코드블록, 머리말, 설명문을 붙이지 마십시오.

{{
  "keywords": [
    {{
      "word": "키워드",
      "frequency": 0,
      "meaning": "의미 설명"
    }}
  ],
  "sentiment_summary": {{
    "positive_count": 0,
    "neutral_count": 0,
    "negative_count": 0,
    "positive_reasons": ["사유"],
    "neutral_reasons": ["사유"],
    "negative_reasons": ["사유"]
  }},
  "improvement_suggestions": [
    "개선방안"
  ],
  "individual_feedback": [
    {{
      "review_no": 1,
      "review_snippet": "원문 일부 50자 이내",
      "feedback": "개인별 피드백"
    }}
  ]
}}
"""


def analyze_course(group, course_title, review_col, rating_col):
    sample = group.copy()

    if len(sample) > MAX_REVIEWS_PER_COURSE:
        sample = sample.sample(
            MAX_REVIEWS_PER_COURSE,
            random_state=42,
        ).sort_index()

    lines = []
    for review_no, (_, row) in enumerate(sample.iterrows(), start=1):
        rating_txt = ""
        if rating_col and pd.notna(row.get(rating_col)):
            rating_txt = f"[별점 {row[rating_col]}] "

        text = str(row[review_col]).replace("\n", " ").strip()
        lines.append(f"[리뷰번호 {review_no}] {rating_txt}{text}")

    prompt = PROMPT_TEMPLATE.format(
        course_title=course_title,
        n_reviews=len(sample),
        reviews_block="\n".join(lines),
    )

    response = client.responses.create(
        model=MODEL,
        input=prompt,
        max_output_tokens=7000,
    )

    result = extract_json(response.output_text)
    result.setdefault("keywords", [])
    result.setdefault("sentiment_summary", {})
    result.setdefault("improvement_suggestions", [])
    result.setdefault("individual_feedback", [])

    return result, len(sample)


st.subheader("1. 분석 데이터")

source = st.radio(
    "데이터 사용 방법",
    ["샘플 데이터로 체험", "내 파일 업로드"],
    horizontal=True,
)

df = None
source_name = ""

if source == "샘플 데이터로 체험":
    if SAMPLE_FILE.exists():
        df = load_sample(str(SAMPLE_FILE))
        source_name = SAMPLE_FILE.name
        st.success(f"샘플 데이터 `{SAMPLE_FILE.name}`를 불러왔습니다.")
    else:
        st.warning(
            "`reviews_seoul_center_final.csv` 파일이 앱 폴더에 없습니다. "
            "GitHub 저장소에 app.py와 같은 위치로 올리거나, 아래에서 직접 업로드해주세요."
        )
        source = "내 파일 업로드"

if source == "내 파일 업로드":
    uploaded_file = st.file_uploader(
        "CSV 또는 Excel(.xlsx) 파일을 업로드하세요.",
        type=["csv", "xlsx"],
    )
    if uploaded_file is not None:
        try:
            df = load_dataframe(uploaded_file)
            source_name = uploaded_file.name
        except Exception as e:
            st.error(f"파일을 읽지 못했습니다: {e}")
            st.stop()

if df is None:
    st.info("샘플 파일을 저장소에 올리거나 분석할 파일을 업로드해주세요.")
    st.stop()


course_col = find_col(
    df,
    ["과정명", "훈련과정명", "훈련과정", "course", "title"],
)
review_col = find_col(
    df,
    ["후기", "리뷰본문", "리뷰", "의견", "review", "content", "text"],
)
rating_col = find_col(
    df,
    ["별점", "평점", "점수", "rating", "score"],
)

c1, c2, c3 = st.columns(3)
c1.metric("과정명 컬럼", course_col or "인식 실패")
c2.metric("후기 컬럼", review_col or "인식 실패")
c3.metric("별점 컬럼", rating_col or "없음")

if course_col is None or review_col is None:
    st.error(
        "과정명 또는 후기 컬럼을 자동으로 찾지 못했습니다. "
        "예: title/과정명, content/후기 형태의 컬럼이 필요합니다."
    )
    st.stop()

df = clean_dataframe(df, course_col, review_col, rating_col)

if df.empty:
    st.error("분석 가능한 후기 데이터가 없습니다.")
    st.stop()


st.subheader("2. 데이터 요약")

m1, m2, m3, m4 = st.columns(4)
m1.metric("후기 수", f"{len(df):,}건")
m2.metric("과정 수", f"{df[course_col].nunique():,}개")

if rating_col and df[rating_col].notna().any():
    m3.metric("평균 별점", f"{df[rating_col].mean():.2f}")
else:
    m3.metric("평균 별점", "-")

m4.metric("분석 모델", "GPT-5.6 Luna")

with st.expander("원본 데이터 미리보기"):
    st.dataframe(df.head(30), use_container_width=True)

if len(df) > MAX_REVIEWS_PER_COURSE:
    st.caption(
        f"※ 과정별 후기가 {MAX_REVIEWS_PER_COURSE}건을 초과하면 "
        f"API 비용과 처리시간을 줄이기 위해 최대 {MAX_REVIEWS_PER_COURSE}건을 표본 분석합니다."
    )


st.subheader("3. AI 분석")

courses = df[course_col].dropna().astype(str).unique().tolist()
selected_courses = st.multiselect(
    "분석할 과정을 선택하세요.",
    options=courses,
    default=courses,
)

if st.button("🚀 AI 분석 시작", type="primary", use_container_width=True):
    if not selected_courses:
        st.warning("분석할 과정을 하나 이상 선택해주세요.")
        st.stop()

    results = {}
    errors = {}
    analyzed_counts = {}

    progress = st.progress(0, text="분석 준비 중...")

    for i, course in enumerate(selected_courses):
        progress.progress(
            i / len(selected_courses),
            text=f"분석 중: {course}",
        )

        group = df[df[course_col].astype(str) == str(course)]

        try:
            result, analyzed_n = analyze_course(
                group=group,
                course_title=course,
                review_col=review_col,
                rating_col=rating_col,
            )
            results[course] = result
            analyzed_counts[course] = analyzed_n
        except Exception as e:
            errors[course] = str(e)

        progress.progress(
            (i + 1) / len(selected_courses),
            text=f"처리 완료: {course}",
        )

    progress.empty()
    st.session_state["results"] = results
    st.session_state["errors"] = errors
    st.session_state["analyzed_counts"] = analyzed_counts
    st.session_state["source_name"] = source_name


if "results" in st.session_state:
    st.subheader("4. 분석 결과")

    results = st.session_state["results"]
    errors = st.session_state.get("errors", {})
    analyzed_counts = st.session_state.get("analyzed_counts", {})

    if errors:
        with st.expander("⚠️ 분석 오류 확인"):
            for course, err in errors.items():
                st.error(f"{course}: {err}")

    for course, result in results.items():
        st.divider()
        st.header(f"📘 {course}")

        sentiment = result.get("sentiment_summary", {})
        pos = int(sentiment.get("positive_count", 0) or 0)
        neu = int(sentiment.get("neutral_count", 0) or 0)
        neg = int(sentiment.get("negative_count", 0) or 0)
        total_sent = pos + neu + neg

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("분석 후기", f"{analyzed_counts.get(course, 0):,}건")

        if total_sent > 0:
            k2.metric("긍정", f"{pos:,}건", f"{pos / total_sent * 100:.1f}%")
            k3.metric("중립", f"{neu:,}건", f"{neu / total_sent * 100:.1f}%")
            k4.metric("부정", f"{neg:,}건", f"{neg / total_sent * 100:.1f}%")
        else:
            k2.metric("긍정", "-")
            k3.metric("중립", "-")
            k4.metric("부정", "-")

        left, right = st.columns([1.1, 0.9])

        with left:
            st.markdown("#### 🔑 핵심 키워드 TOP 10")
            keywords = result.get("keywords", [])[:10]

            if keywords:
                keyword_rows = []
                for kw in keywords:
                    keyword_rows.append(
                        {
                            "키워드": kw.get("word", ""),
                            "빈도": int(kw.get("frequency", 0) or 0),
                            "의미": kw.get("meaning", ""),
                        }
                    )

                keyword_df = pd.DataFrame(keyword_rows)
                st.dataframe(
                    keyword_df,
                    use_container_width=True,
                    hide_index=True,
                )

                if keyword_df["빈도"].sum() > 0:
                    st.bar_chart(
                        keyword_df.set_index("키워드")["빈도"],
                        horizontal=True,
                    )
            else:
                st.info("키워드 결과가 없습니다.")

        with right:
            st.markdown("#### 😊 감성분포")
            sentiment_df = pd.DataFrame(
                {
                    "감성": ["긍정", "중립", "부정"],
                    "건수": [pos, neu, neg],
                }
            )
            st.bar_chart(sentiment_df.set_index("감성"))

        reason_cols = st.columns(3)

        with reason_cols[0]:
            st.markdown("##### 긍정 주요 사유")
            positive_reasons = safe_list(sentiment, "positive_reasons", 5)
            if positive_reasons:
                for x in positive_reasons:
                    st.markdown(f"- {x}")
            else:
                st.caption("해당 의견 없음")

        with reason_cols[1]:
            st.markdown("##### 중립 주요 사유")
            neutral_reasons = safe_list(sentiment, "neutral_reasons", 5)
            if neutral_reasons:
                for x in neutral_reasons:
                    st.markdown(f"- {x}")
            else:
                st.caption("해당 의견 없음")

        with reason_cols[2]:
            st.markdown("##### 부정·개선 요구 주요 사유")
            negative_reasons = safe_list(sentiment, "negative_reasons", 5)
            if negative_reasons:
                for x in negative_reasons:
                    st.markdown(f"- {x}")
            else:
                st.caption("해당 의견 없음")

        st.markdown("#### 🛠️ 개선 방안 5가지")
        improvements = result.get("improvement_suggestions", [])[:5]

        if improvements:
            for idx, imp in enumerate(improvements, start=1):
                st.markdown(f"**{idx}.** {imp}")
        else:
            st.info("개선방안 결과가 없습니다.")

        st.markdown("#### 💬 개인별 피드백")
        feedback_list = result.get("individual_feedback", [])[:10]

        if feedback_list:
            for fb in feedback_list:
                review_no = fb.get("review_no", "")
                snippet = fb.get("review_snippet", "")
                title = f"리뷰 {review_no} · {snippet}" if review_no else snippet

                with st.expander(title or "개인별 피드백"):
                    st.write(fb.get("feedback", ""))
        else:
            st.info("개인별 피드백 결과가 없습니다.")

    st.divider()

    st.download_button(
        "📥 전체 분석 결과 JSON 다운로드",
        data=json.dumps(results, ensure_ascii=False, indent=2),
        file_name="analysis_results.json",
        mime="application/json",
        use_container_width=True,
    )

    st.caption(
        "본 결과는 생성형 AI를 활용한 자동 분석 결과이며, "
        "실제 업무 적용 시 담당자의 최종 검토를 권장합니다."
    )
