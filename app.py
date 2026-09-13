import io
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import streamlit as st
from openai import OpenAI

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# =========================================================
# 0. 기본 설정
# =========================================================
st.set_page_config(
    page_title="수강후기 AI 분석 리포트",
    page_icon="📊",
    layout="wide",
)

APP_TITLE = "수강후기 AI 분석 리포트"
MODEL = "gpt-5.6-luna"
MAX_PARALLEL_WORKERS = 3
SAMPLE_FILE = Path(__file__).with_name("reviews_seoul_center_final.csv")

st.title(f"📊 {APP_TITLE}")
st.caption(
    "수강후기 데이터를 기반으로 과정별 핵심 키워드, 감성, 주요 의견, "
    "개선방안 및 개인별 피드백을 자동으로 분석합니다."
)


# =========================================================
# 1. OpenAI API
# =========================================================
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error(
        "OpenAI API 키가 설정되지 않았습니다. "
        "Streamlit Cloud → App settings → Secrets에 "
        'OPENAI_API_KEY = "sk-..." 형식으로 등록해주세요.'
    )
    st.stop()


# =========================================================
# 2. 데이터 유틸리티
# =========================================================
def normalize_name(value):
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )


def find_col(df, exact_candidates, partial_candidates=None, exclude=None):
    partial_candidates = partial_candidates or []
    exclude = {normalize_name(x) for x in (exclude or [])}
    normalized = {col: normalize_name(col) for col in df.columns}

    for candidate in exact_candidates:
        c = normalize_name(candidate)
        for col, ncol in normalized.items():
            if ncol in exclude:
                continue
            if ncol == c:
                return col

    for candidate in partial_candidates:
        c = normalize_name(candidate)
        for col, ncol in normalized.items():
            if ncol in exclude:
                continue
            if c in ncol or ncol in c:
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
    original_n = len(work)

    work = work.dropna(subset=[course_col, review_col])
    work[course_col] = work[course_col].astype(str).str.strip()
    work[review_col] = work[review_col].astype(str).str.strip()

    # 의미 없는 공백만 제거하고, 짧은 후기 자체는 보존
    work = work[work[review_col].str.len() >= 1]

    if rating_col:
        work[rating_col] = pd.to_numeric(work[rating_col], errors="coerce")

    removed_n = original_n - len(work)
    return work.reset_index(drop=True), removed_n


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
        raise ValueError("모델 응답에서 JSON 객체를 찾지 못했습니다.")

    return json.loads(text[start:end + 1])


def safe_list(obj, key, limit=None):
    value = obj.get(key, [])
    if not isinstance(value, list):
        return []
    return value[:limit] if limit else value


def compact_text(text):
    return re.sub(r"\s+", "", str(text).lower())


def compute_keyword_frequencies(sample, review_col, keywords):
    review_texts = [compact_text(x) for x in sample[review_col].astype(str)]

    for kw in keywords:
        terms = [kw.get("word", "")]
        aliases = kw.get("aliases", [])

        if isinstance(aliases, list):
            terms.extend(aliases)

        terms = [compact_text(x) for x in terms if str(x).strip()]
        terms = list(dict.fromkeys(terms))

        count = 0
        for text in review_texts:
            if any(term and term in text for term in terms):
                count += 1

        kw["frequency"] = count

    return keywords


def build_course_overview(df, course_col, rating_col):
    if rating_col and df[rating_col].notna().any():
        overview = (
            df.groupby(course_col)
            .agg(
                후기수=(course_col, "size"),
                평균별점=(rating_col, "mean"),
            )
            .reset_index()
            .rename(columns={course_col: "과정명"})
        )
        overview["평균별점"] = overview["평균별점"].round(2)
    else:
        overview = (
            df.groupby(course_col)
            .size()
            .reset_index(name="후기수")
            .rename(columns={course_col: "과정명"})
        )
        overview["평균별점"] = None

    return overview.sort_values("후기수", ascending=False).reset_index(drop=True)


# =========================================================
# 3. 리포트 생성
# =========================================================
def result_to_rows(result):
    counts = result.get("sentiment_counts", {})
    return {
        "positive": int(counts.get("positive", 0) or 0),
        "neutral": int(counts.get("neutral", 0) or 0),
        "negative": int(counts.get("negative", 0) or 0),
        "analysis_n": int(result.get("analysis_review_count", 0) or 0),
    }



def _keyword_terms(result):
    terms = []
    for kw in result.get("keywords", [])[:10]:
        group = [kw.get("word", "")]
        aliases = kw.get("aliases", [])
        if isinstance(aliases, list):
            group.extend(aliases)
        group = [compact_text(x) for x in group if str(x).strip()]
        if group:
            terms.append(group)
    return terms


def _review_information_score(text, keyword_groups):
    raw = str(text).strip()
    compact = compact_text(raw)

    # 너무 긴 글만 무조건 뽑히지 않도록 길이 점수는 300자에서 포화
    length_score = min(len(raw), 300) / 300

    covered = set()
    for i, group in enumerate(keyword_groups):
        if any(term and term in compact for term in group):
            covered.add(i)

    keyword_score = min(len(covered), 4) / 4

    cue_words = [
        "좋", "도움", "유익", "추천", "만족",
        "아쉽", "부족", "어렵", "불편", "개선",
        "프로젝트", "강사", "취업", "실습", "장비",
        "커리큘럼", "난이도", "시간", "환경",
    ]
    cue_hits = sum(1 for cue in cue_words if cue in raw)
    cue_score = min(cue_hits, 5) / 5

    score = 0.45 * length_score + 0.40 * keyword_score + 0.15 * cue_score
    return score, covered


def select_representative_reviews(sample, result, review_col, id_col, target_n=10):
    """
    대표 후기 선정:
    1) 긍정/중립/부정 층화(기본 4/2/4)
    2) 후기 구체성/정보량
    3) 주요 키워드·이슈 다양성
    4) 부정·개선요구 의견을 약간 우대
    """
    work = sample.copy()

    if id_col:
        work["_rid"] = work[id_col].astype(str)
    else:
        work["_rid"] = (work.index + 1).astype(str)

    if len(work) <= target_n:
        work["_sentiment"] = "neutral"
        sentiment_map = {
            str(item.get("review_id", "")): str(item.get("sentiment", "")).lower()
            for item in result.get("review_sentiments", [])
        }
        work["_sentiment"] = work["_rid"].map(sentiment_map).fillna("neutral")
        work["_selection_reason"] = "전체 후기"
        return work

    sentiment_map = {
        str(item.get("review_id", "")): str(item.get("sentiment", "")).lower()
        for item in result.get("review_sentiments", [])
        if str(item.get("sentiment", "")).lower()
        in {"positive", "neutral", "negative"}
    }
    work["_sentiment"] = work["_rid"].map(sentiment_map).fillna("neutral")

    keyword_groups = _keyword_terms(result)
    score_meta = work[review_col].apply(
        lambda x: _review_information_score(x, keyword_groups)
    )
    work["_info_score"] = score_meta.apply(lambda x: x[0])
    work["_covered_keywords"] = score_meta.apply(lambda x: x[1])

    quotas = {"positive": 4, "neutral": 2, "negative": 4}
    selected_indices = []
    globally_covered = set()

    # 개선 이슈가 묻히지 않도록 부정 -> 중립 -> 긍정 순으로 대표성 확보
    for sentiment in ["negative", "neutral", "positive"]:
        candidates = work[work["_sentiment"] == sentiment].copy()
        quota = min(quotas[sentiment], len(candidates))

        for _ in range(quota):
            if candidates.empty:
                break

            candidates["_diversity_bonus"] = candidates["_covered_keywords"].apply(
                lambda s: len(set(s) - globally_covered) * 0.12
            )
            candidates["_rank_score"] = (
                candidates["_info_score"] + candidates["_diversity_bonus"]
            )

            chosen_idx = candidates["_rank_score"].idxmax()
            selected_indices.append(chosen_idx)
            globally_covered |= set(candidates.loc[chosen_idx, "_covered_keywords"])
            candidates = candidates.drop(index=chosen_idx)

    # 특정 감성의 후기가 부족하면 남는 자리를 전체 후보에서 채운다.
    remaining = target_n - len(selected_indices)
    candidates = work.drop(index=selected_indices, errors="ignore").copy()

    for _ in range(min(remaining, len(candidates))):
        candidates["_diversity_bonus"] = candidates["_covered_keywords"].apply(
            lambda s: len(set(s) - globally_covered) * 0.12
        )
        candidates["_sentiment_bonus"] = candidates["_sentiment"].map(
            {"negative": 0.10, "neutral": 0.05, "positive": 0.0}
        )
        candidates["_rank_score"] = (
            candidates["_info_score"]
            + candidates["_diversity_bonus"]
            + candidates["_sentiment_bonus"]
        )

        chosen_idx = candidates["_rank_score"].idxmax()
        selected_indices.append(chosen_idx)
        globally_covered |= set(candidates.loc[chosen_idx, "_covered_keywords"])
        candidates = candidates.drop(index=chosen_idx)

    selected = work.loc[selected_indices].copy()
    selected["_selection_reason"] = selected["_sentiment"].map(
        {
            "positive": "긍정 대표 의견",
            "neutral": "중립·혼합 대표 의견",
            "negative": "부정·개선요구 대표 의견",
        }
    )
    return selected


FEEDBACK_PROMPT_TEMPLATE = """
당신은 직업훈련기관의 교육품질 개선 담당자입니다.

아래 대표 수강후기는 단순 무작위가 아니라
긍정·중립·부정 의견의 균형, 후기의 구체성,
주요 키워드와 이슈의 다양성을 고려하여 선정되었습니다.

각 후기 원문을 충분히 반영하여 해당 교육생에게 전달할 수 있는
정중하고 구체적인 피드백을 작성하십시오.

규칙:
- review_id는 제공된 실제 ID를 그대로 사용
- feedback은 2~4문장
- 단순 감사 인사에 그치지 말고 의견을 어떻게 유지·보완·개선할지 드러낼 것
- 후기에서 확인되지 않는 사실을 만들지 말 것
- 개인정보를 추정하지 말 것
- 반드시 JSON만 반환

[대표 후기]
{reviews_block}

{{
  "individual_feedback": [
    {{
      "review_id": "실제 리뷰ID",
      "feedback": "개인별 피드백"
    }}
  ]
}}
"""


def generate_feedback_for_selected(selected, review_col, id_col):
    lines = []

    for _, row in selected.iterrows():
        rid = str(row[id_col]) if id_col else str(row["_rid"])
        sentiment = str(row["_sentiment"])
        reason = str(row["_selection_reason"])
        original = str(row[review_col]).replace("\n", " ").strip()

        lines.append(
            f"[리뷰ID {rid}] [감성 {sentiment}] [선정사유 {reason}] "
            f"[원문] {original}"
        )

    prompt = FEEDBACK_PROMPT_TEMPLATE.format(
        reviews_block="\n".join(lines)
    )

    response = client.responses.create(
        model=MODEL,
        input=prompt,
        max_output_tokens=5000,
    )

    payload = extract_json(response.output_text)
    feedbacks = payload.get("individual_feedback", [])

    original_lookup = {}
    meta_lookup = {}

    for _, row in selected.iterrows():
        rid = str(row[id_col]) if id_col else str(row["_rid"])
        original_lookup[rid] = str(row[review_col]).strip()
        meta_lookup[rid] = {
            "sentiment": str(row["_sentiment"]),
            "selection_reason": str(row["_selection_reason"]),
        }

    cleaned = []
    seen = set()

    for item in feedbacks:
        rid = str(item.get("review_id", ""))
        if rid in original_lookup and rid not in seen:
            cleaned.append(
                {
                    "review_id": rid,
                    "sentiment": meta_lookup[rid]["sentiment"],
                    "selection_reason": meta_lookup[rid]["selection_reason"],
                    "original_review": original_lookup[rid],
                    "feedback": str(item.get("feedback", "")).strip(),
                }
            )
            seen.add(rid)

    # 모델이 일부 항목을 누락해도 대표 후기 자체는 화면과 Word에 남긴다.
    for rid, original in original_lookup.items():
        if rid not in seen:
            cleaned.append(
                {
                    "review_id": rid,
                    "sentiment": meta_lookup[rid]["sentiment"],
                    "selection_reason": meta_lookup[rid]["selection_reason"],
                    "original_review": original,
                    "feedback": (
                        "대표 의견으로 선정되었습니다. "
                        "해당 의견을 교육과정 운영 개선 시 검토하겠습니다."
                    ),
                }
            )

    return cleaned[:10]


def build_docx_report(results):
    doc = Document()

    styles = doc.styles
    styles["Normal"].font.name = "Malgun Gothic"
    styles["Normal"].font.size = Pt(10)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("수강후기 AI 분석 리포트")
    run.bold = True
    run.font.size = Pt(18)

    doc.add_paragraph(
        "생성형 AI를 활용하여 과정별 핵심 키워드, 감성, 주요 의견, "
        "개선방안 및 개인별 피드백을 자동 분석한 결과입니다."
    )

    for idx, (course, result) in enumerate(results.items(), start=1):
        doc.add_heading(f"{idx}. {course}", level=1)

        summary = result.get("executive_summary", "")
        if summary:
            doc.add_heading("AI 핵심 요약", level=2)
            doc.add_paragraph(summary)

        counts = result_to_rows(result)
        doc.add_heading("감성 분석", level=2)

        table = doc.add_table(rows=2, cols=4)
        table.style = "Table Grid"
        headers = ["분석 후기", "긍정", "중립", "부정"]
        values = [
            counts["analysis_n"],
            counts["positive"],
            counts["neutral"],
            counts["negative"],
        ]
        for j, h in enumerate(headers):
            table.cell(0, j).text = h
            table.cell(1, j).text = str(values[j])

        doc.add_heading("핵심 키워드 TOP 10", level=2)
        kw_table = doc.add_table(rows=1, cols=3)
        kw_table.style = "Table Grid"
        kw_table.rows[0].cells[0].text = "키워드"
        kw_table.rows[0].cells[1].text = "빈도"
        kw_table.rows[0].cells[2].text = "의미"

        for kw in result.get("keywords", [])[:10]:
            cells = kw_table.add_row().cells
            cells[0].text = str(kw.get("word", ""))
            cells[1].text = str(kw.get("frequency", 0))
            cells[2].text = str(kw.get("meaning", ""))

        reasons = result.get("sentiment_reasons", {})
        for label, key in [
            ("긍정 주요 사유", "positive"),
            ("중립 주요 사유", "neutral"),
            ("부정·개선 요구 주요 사유", "negative"),
        ]:
            doc.add_heading(label, level=2)
            vals = safe_list(reasons, key, 5)
            if vals:
                for v in vals:
                    doc.add_paragraph(str(v), style="List Bullet")
            else:
                doc.add_paragraph("해당 의견 없음")

        doc.add_heading("개선 방안 5가지", level=2)
        for i, item in enumerate(
            result.get("improvement_suggestions", [])[:5],
            start=1,
        ):
            doc.add_paragraph(f"{i}. {item}")

        doc.add_heading("개인별 피드백", level=2)
        for fb in result.get("individual_feedback", [])[:10]:
            rid = fb.get("review_id", "")
            sentiment = fb.get("sentiment", "")
            selection_reason = fb.get("selection_reason", "")
            original_review = fb.get("original_review", "")
            feedback = fb.get("feedback", "")

            sentiment_label = {
                "positive": "긍정",
                "neutral": "중립",
                "negative": "부정",
            }.get(sentiment, sentiment)

            doc.add_paragraph(
                f"교육생 의견 #{rid} ({sentiment_label}) - {selection_reason}",
                style="List Bullet",
            )
            doc.add_paragraph(f"원문: {original_review}")
            doc.add_paragraph(f"AI 피드백: {feedback}")

        if idx < len(results):
            doc.add_page_break()

    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run(
        "본 결과는 생성형 AI 기반 자동 분석 결과이며, 실제 업무 적용 시 담당자 최종 검토를 권장합니다."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# 4. 분석 프롬프트
# =========================================================
PROMPT_TEMPLATE = """
당신은 직업훈련기관의 강의평가·수강후기를 분석하는 데이터 분석가입니다.

아래는 하나의 교육과정에 대한 후기 목록입니다.
각 행에는 [리뷰ID], 선택적으로 [별점], [후기본문]이 있습니다.

[과정명]
{course_title}

[분석대상 후기 수]
{n_reviews}

[후기 목록]
{reviews_block}

다음 업무를 수행하십시오.

1. executive_summary
- 이 과정의 강점과 핵심 개선과제를 2~3문장으로 요약
- 후기에서 확인되는 내용만 사용

2. keywords
- 교육과정 개선에 의미 있는 핵심 키워드 정확히 10개
- word: 대표 키워드
- aliases: 후기에서 실제 함께 사용되는 유사 표현 0~5개
- meaning: 이 과정에서 해당 키워드가 의미하는 바 한 문장
- frequency는 반환하지 마십시오. Python이 실제 후기에서 다시 계산합니다.

3. review_sentiments
- 모든 {n_reviews}개 후기 각각을 positive / neutral / negative 중 하나로 분류
- review_id는 제공된 실제 [리뷰ID]를 그대로 사용
- 후기 수와 sentiment label 수가 정확히 일치해야 함

4. sentiment_reasons
- positive / neutral / negative의 대표적인 사유를 각각 최대 5개
- 해당 감성이 거의 없으면 빈 배열 허용
- 억지로 내용을 만들어내지 말 것

5. improvement_suggestions
- 후기 내용에 직접 근거한 구체적이고 실행 가능한 개선방안 정확히 5개

중요:
- 후기에서 확인되지 않는 사실을 만들지 마십시오.
- 개인정보를 추정하거나 생성하지 마십시오.
- 리뷰 ID와 후기 내용을 혼동하지 마십시오.
- 반드시 아래 JSON 구조만 반환하십시오.
- 코드블록, 머리말, 설명문을 붙이지 마십시오.

{{
  "executive_summary": "요약",
  "keywords": [
    {{
      "word": "키워드",
      "aliases": ["유사표현1", "유사표현2"],
      "meaning": "의미"
    }}
  ],
  "review_sentiments": [
    {{
      "review_id": "실제 리뷰ID",
      "sentiment": "positive"
    }}
  ],
  "sentiment_reasons": {{
    "positive": ["사유"],
    "neutral": ["사유"],
    "negative": ["사유"]
  }},
  "improvement_suggestions": [
    "개선방안"
  ]}}
"""


def analyze_course(group, course_title, review_col, rating_col, id_col):
    # 모든 유효 후기 사용
    sample = group.copy()

    if id_col:
        sample["_analysis_review_id"] = sample[id_col].astype(str)
    else:
        sample["_analysis_review_id"] = (sample.index + 1).astype(str)

    lines = []

    for _, row in sample.iterrows():
        review_id = row["_analysis_review_id"]

        rating_txt = ""
        if rating_col and pd.notna(row.get(rating_col)):
            rating_txt = f"[별점 {row[rating_col]}] "

        text = str(row[review_col]).replace("\n", " ").strip()

        lines.append(
            f"[리뷰ID {review_id}] {rating_txt}[후기본문] {text}"
        )

    prompt = PROMPT_TEMPLATE.format(
        course_title=course_title,
        n_reviews=len(sample),
        reviews_block="\n".join(lines),
    )

    response = client.responses.create(
        model=MODEL,
        input=prompt,
        max_output_tokens=14000,
    )

    result = extract_json(response.output_text)

    result.setdefault("executive_summary", "")
    result.setdefault("keywords", [])
    result.setdefault("review_sentiments", [])
    result.setdefault(
        "sentiment_reasons",
        {"positive": [], "neutral": [], "negative": []},
    )
    result.setdefault("improvement_suggestions", [])
    result.setdefault("individual_feedback", [])

    result["keywords"] = compute_keyword_frequencies(
        sample,
        review_col,
        result.get("keywords", [])[:10],
    )

    valid_ids = set(sample["_analysis_review_id"].astype(str))
    sentiment_map = {}

    for item in result.get("review_sentiments", []):
        rid = str(item.get("review_id", ""))
        sentiment = str(item.get("sentiment", "")).lower()

        if rid in valid_ids and sentiment in {
            "positive",
            "neutral",
            "negative",
        }:
            sentiment_map[rid] = sentiment

    counts = {
        "positive": sum(v == "positive" for v in sentiment_map.values()),
        "neutral": sum(v == "neutral" for v in sentiment_map.values()),
        "negative": sum(v == "negative" for v in sentiment_map.values()),
    }

    result["sentiment_counts"] = counts
    result["classified_reviews"] = len(sentiment_map)
    result["analysis_review_count"] = len(sample)

    representative = select_representative_reviews(
        sample=sample,
        result=result,
        review_col=review_col,
        id_col=id_col,
        target_n=10,
    )

    result["individual_feedback"] = generate_feedback_for_selected(
        selected=representative,
        review_col=review_col,
        id_col=id_col,
    )

    return result


# =========================================================
# 5. 데이터 선택
# =========================================================
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
            "GitHub 저장소에 app.py와 같은 위치로 올리거나 직접 업로드해주세요."
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


# =========================================================
# 6. 컬럼 인식
# =========================================================
course_col = find_col(
    df,
    exact_candidates=["title", "과정명", "훈련과정명", "course"],
    partial_candidates=["훈련과정"],
)

review_col = find_col(
    df,
    exact_candidates=[
        "content",
        "review_text",
        "reviewtext",
        "후기",
        "리뷰본문",
        "리뷰내용",
        "의견",
        "text",
    ],
    partial_candidates=["후기내용", "수강후기", "교육후기"],
    exclude=["review_id", "reviewid", "id"],
)

rating_col = find_col(
    df,
    exact_candidates=["rating", "별점", "평점", "score", "점수"],
)

id_col = find_col(
    df,
    exact_candidates=["review_id", "reviewid", "리뷰id", "후기id"],
)

c1, c2, c3, c4 = st.columns(4)
c1.metric("과정명 컬럼", course_col or "인식 실패")
c2.metric("후기본문 컬럼", review_col or "인식 실패")
c3.metric("별점 컬럼", rating_col or "없음")
c4.metric("리뷰ID 컬럼", id_col or "없음")

if course_col is None or review_col is None:
    st.error(
        "과정명 또는 후기본문 컬럼을 자동으로 찾지 못했습니다. "
        "예: title/과정명, content/후기 형태의 컬럼이 필요합니다."
    )
    st.stop()

df, removed_n = clean_dataframe(df, course_col, review_col, rating_col)

if df.empty:
    st.error("분석 가능한 후기 데이터가 없습니다.")
    st.stop()


# =========================================================
# 7. 전체 데이터 대시보드
# =========================================================
st.subheader("2. 데이터 요약")

overview = build_course_overview(df, course_col, rating_col)

m1, m2, m3, m4 = st.columns(4)
m1.metric("전체 후기", f"{len(df):,}건")
m2.metric("과정 수", f"{df[course_col].nunique():,}개")

if rating_col and df[rating_col].notna().any():
    m3.metric("전체 평균 별점", f"{df[rating_col].mean():.2f}")
else:
    m3.metric("전체 평균 별점", "-")

m4.metric("분석 모델", "GPT-5.6 Luna")

if removed_n > 0:
    st.caption(
        f"※ 과정명 또는 후기본문이 비어 있는 {removed_n:,}건은 분석에서 제외했습니다."
    )

left_overview, right_overview = st.columns([1.15, 0.85])

with left_overview:
    st.markdown("#### 과정별 현황")
    st.dataframe(
        overview,
        use_container_width=True,
        hide_index=True,
    )

with right_overview:
    st.markdown("#### 과정별 후기 수")
    course_count_chart = overview[["과정명", "후기수"]].set_index("과정명")
    st.bar_chart(course_count_chart, horizontal=True)

if rating_col and df[rating_col].notna().any():
    st.markdown("#### 별점 분포")

    rating_dist = (
        df[rating_col]
        .dropna()
        .round(1)
        .value_counts()
        .sort_index()
        .rename_axis("별점")
        .reset_index(name="후기수")
    )

    st.bar_chart(
        rating_dist.set_index("별점")["후기수"],
        use_container_width=True,
    )

with st.expander("원본 데이터 미리보기"):
    display_cols = [x for x in [id_col, course_col, rating_col, review_col] if x]
    st.dataframe(
        df[display_cols].head(30),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 8. 분석 실행
# =========================================================
st.subheader("3. AI 분석")

courses = df[course_col].dropna().astype(str).unique().tolist()

selected_courses = st.multiselect(
    "분석할 과정을 선택하세요.",
    options=courses,
    default=courses,
)

st.caption(
    f"※ 여러 과정을 선택하면 최대 {MAX_PARALLEL_WORKERS}개 과정을 동시에 분석합니다. "
    "같은 브라우저 세션에서 이미 분석한 과정은 재사용합니다."
)

if st.button("🚀 AI 분석 시작", type="primary", use_container_width=True):
    if not selected_courses:
        st.warning("분석할 과정을 하나 이상 선택해주세요.")
        st.stop()

    # 같은 브라우저 세션에서 이미 분석한 과정은 재사용
    previous_results = st.session_state.get("results", {})
    previous_errors = st.session_state.get("errors", {})

    results = dict(previous_results)
    errors = dict(previous_errors)

    # 현재 선택된 과정 중 아직 결과가 없는 과정만 새로 분석
    courses_to_run = [
        course for course in selected_courses
        if course not in results
    ]

    # 선택에서 빠진 과정은 화면 결과에서도 제거
    results = {
        course: result
        for course, result in results.items()
        if course in selected_courses
    }
    errors = {
        course: err
        for course, err in errors.items()
        if course in selected_courses
    }

    progress_bar = st.progress(0)
    status_box = st.empty()

    if not courses_to_run:
        status_box.success(
            "✅ 선택한 과정은 현재 세션에 이미 분석 결과가 있어 재사용했습니다."
        )
        progress_bar.progress(1.0)

    else:
        worker_count = min(
            MAX_PARALLEL_WORKERS,
            len(courses_to_run),
        )

        status_box.info(
            f"🔄 최대 {worker_count}개 과정을 병렬 분석합니다. "
            f"신규 분석 대상 {len(courses_to_run)}개"
        )

        completed = 0

        def _run_one_course(course_name):
            group = df[
                df[course_col].astype(str) == str(course_name)
            ].copy()

            result = analyze_course(
                group=group,
                course_title=course_name,
                review_col=review_col,
                rating_col=rating_col,
                id_col=id_col,
            )
            return course_name, result, len(group)

        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            future_map = {
                executor.submit(_run_one_course, course): course
                for course in courses_to_run
            }

            for future in as_completed(future_map):
                course = future_map[future]

                try:
                    course_name, result, review_count = future.result()
                    results[course_name] = result
                    errors.pop(course_name, None)

                    completed += 1
                    status_box.info(
                        f"🔄 {completed}/{len(courses_to_run)} 신규 과정 완료 · "
                        f"{course_name} ({review_count:,}건)"
                    )
                except Exception as e:
                    errors[course] = str(e)
                    completed += 1
                    status_box.warning(
                        f"⚠️ {completed}/{len(courses_to_run)} 처리 · "
                        f"{course} 분석 실패"
                    )

                progress_bar.progress(
                    completed / len(courses_to_run)
                )

        status_box.success(
            f"✅ 분석 완료 · 신규 {len(courses_to_run)}개 과정 처리 · "
            f"최대 {worker_count}개 병렬 실행"
        )

    time.sleep(0.2)

    st.session_state["results"] = results
    st.session_state["errors"] = errors
    st.session_state["source_name"] = source_name


# =========================================================
# 9. 결과 렌더링
# =========================================================
if "results" in st.session_state:
    st.subheader("4. 분석 결과")

    results = st.session_state["results"]
    errors = st.session_state.get("errors", {})

    if errors:
        with st.expander("⚠️ 분석 오류 확인"):
            for course, err in errors.items():
                st.error(f"{course}: {err}")

    for course, result in results.items():
        st.divider()
        st.header(f"📘 {course}")

        summary = result.get("executive_summary", "")
        if summary:
            st.info(f"**AI 핵심 요약**\n\n{summary}")

        counts = result.get(
            "sentiment_counts",
            {"positive": 0, "neutral": 0, "negative": 0},
        )

        pos = int(counts.get("positive", 0) or 0)
        neu = int(counts.get("neutral", 0) or 0)
        neg = int(counts.get("negative", 0) or 0)

        classified = int(result.get("classified_reviews", 0) or 0)
        analysis_n = int(result.get("analysis_review_count", 0) or 0)

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("분석 후기", f"{analysis_n:,}건")

        if classified > 0:
            k2.metric("긍정", f"{pos:,}건", f"{pos / classified * 100:.1f}%")
            k3.metric("중립", f"{neu:,}건", f"{neu / classified * 100:.1f}%")
            k4.metric("부정", f"{neg:,}건", f"{neg / classified * 100:.1f}%")
        else:
            k2.metric("긍정", "-")
            k3.metric("중립", "-")
            k4.metric("부정", "-")

        if classified != analysis_n:
            st.caption(
                f"※ 감성 분류 응답 {classified:,}건 / 분석 대상 {analysis_n:,}건. "
                "모델 응답 누락이 있을 수 있으므로 최종 업무 적용 시 확인이 필요합니다."
            )

        left, right = st.columns([1.15, 0.85])

        with left:
            st.markdown("#### 🔑 핵심 키워드 TOP 10")

            keyword_rows = []

            for kw in result.get("keywords", [])[:10]:
                keyword_rows.append(
                    {
                        "키워드": kw.get("word", ""),
                        "빈도": int(kw.get("frequency", 0) or 0),
                        "의미": kw.get("meaning", ""),
                    }
                )

            if keyword_rows:
                keyword_df = pd.DataFrame(keyword_rows)

                st.dataframe(
                    keyword_df,
                    use_container_width=True,
                    hide_index=True,
                )

                st.bar_chart(
                    keyword_df.set_index("키워드")["빈도"],
                    horizontal=True,
                )
            else:
                st.info("키워드 분석 결과가 없습니다.")

        with right:
            st.markdown("#### 😊 감성 분포")

            sentiment_df = pd.DataFrame(
                {
                    "감성": ["긍정", "중립", "부정"],
                    "건수": [pos, neu, neg],
                }
            )

            st.bar_chart(
                sentiment_df.set_index("감성")["건수"],
                use_container_width=True,
            )

        reasons = result.get("sentiment_reasons", {})
        reason_cols = st.columns(3)

        with reason_cols[0]:
            st.markdown("##### 👍 긍정 주요 사유")
            values = safe_list(reasons, "positive", 5)
            if values:
                for item in values:
                    st.markdown(f"- {item}")
            else:
                st.caption("해당 의견 없음")

        with reason_cols[1]:
            st.markdown("##### ➖ 중립 주요 사유")
            values = safe_list(reasons, "neutral", 5)
            if values:
                for item in values:
                    st.markdown(f"- {item}")
            else:
                st.caption("해당 의견 없음")

        with reason_cols[2]:
            st.markdown("##### 👎 부정·개선 요구 주요 사유")
            values = safe_list(reasons, "negative", 5)
            if values:
                for item in values:
                    st.markdown(f"- {item}")
            else:
                st.caption("해당 의견 없음")

        st.markdown("#### 🛠️ 개선 방안 5가지")

        improvements = result.get("improvement_suggestions", [])[:5]
        if improvements:
            for idx, improvement in enumerate(improvements, start=1):
                st.markdown(f"**{idx}.** {improvement}")
        else:
            st.info("개선방안 결과가 없습니다.")

        st.markdown("#### 💬 대표 교육생 의견 및 개인별 피드백")
        st.caption(
            "긍정·중립·부정 의견의 균형, 후기의 구체성, 주요 키워드·이슈 다양성을 "
            "고려하여 대표 의견 최대 10건을 선정합니다."
        )

        feedbacks = result.get("individual_feedback", [])[:10]

        if feedbacks:
            for fb in feedbacks:
                review_id = fb.get("review_id", "")
                sentiment = fb.get("sentiment", "")
                selection_reason = fb.get("selection_reason", "")
                original_review = fb.get("original_review", "")

                sentiment_label = {
                    "positive": "긍정",
                    "neutral": "중립",
                    "negative": "부정",
                }.get(sentiment, sentiment)

                title = (
                    f"교육생 의견 #{review_id} · {sentiment_label} · {selection_reason}"
                    if review_id
                    else "대표 교육생 의견"
                )

                with st.expander(title):
                    st.markdown("**원문 전체**")
                    st.write(original_review)
                    st.markdown("**AI 피드백**")
                    st.write(fb.get("feedback", ""))
        else:
            st.info("개인별 피드백 결과가 없습니다.")

    st.divider()
    st.markdown("### 📄 분석 리포트 다운로드")

    download_cols = st.columns(2)

    with download_cols[0]:
        st.download_button(
            "📝 Word 리포트 다운로드",
            data=build_docx_report(results),
            file_name="수강후기_AI_분석_리포트.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with download_cols[1]:
        st.download_button(
            "🧾 JSON 원본 다운로드",
            data=json.dumps(results, ensure_ascii=False, indent=2),
            file_name="analysis_results.json",
            mime="application/json",
            use_container_width=True,
        )

    st.caption(
        "본 결과는 생성형 AI를 활용한 자동 분석 결과이며, "
        "실제 업무 적용 시 담당자의 최종 검토를 권장합니다."
    )
