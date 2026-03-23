from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

ANSWER_TEXT_COLS = ["I", "K", "M", "O", "Q", "S"]
ANSWER_NEXT_COLS = ["J", "L", "N", "P", "R", "T"]
ALT_ANSWER_TEXT_KEYS = ["a1", "a2", "a3", "a4", "a5", "a6"]
ALT_ANSWER_NEXT_KEYS = ["nach a1", "nach a2", "nach a3", "nach a4", "nach a5", "nach a6"]


@dataclass
class AnswerOption:
    text: str
    next_question: Optional[str] = None


@dataclass
class Question:
    number: str
    block: str
    text: str
    answer_count: int
    input_type: str
    choice_type: str
    options: List[AnswerOption] = field(default_factory=list)
    source_sheet: str = ""


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def normalize_key(value: Any) -> str:
    return normalize_text(value).lower()


def normalize_question_no(value: Any) -> str:
    return normalize_text(value)


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    for idx in range(min(len(df), 30)):
        row_values = [normalize_key(v) for v in df.iloc[idx].tolist()]
        if "frage nr." in row_values and "frage" in row_values:
            return idx
    return None


def detect_best_sheet(excel_file: pd.ExcelFile) -> Tuple[str, pd.DataFrame, int]:
    best: Optional[Tuple[str, pd.DataFrame, int, int]] = None

    for sheet_name in excel_file.sheet_names:
        raw_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        header_row = find_header_row(raw_df)
        if header_row is None:
            continue

        headers = [normalize_text(v) or f"col_{i}" for i, v in enumerate(raw_df.iloc[header_row].tolist())]
        data_df = raw_df.iloc[header_row + 1 :].copy().reset_index(drop=True)
        data_df.columns = headers

        score = 0
        normalized_columns = {normalize_key(c) for c in data_df.columns}
        if "antwortart" in normalized_columns:
            score += 5
        if "anzahl der antworten" in normalized_columns:
            score += 3
        if "a1" in normalized_columns or "nach a1" in normalized_columns:
            score += 5
        if len(data_df.dropna(how="all")) > 0:
            score += 1

        if best is None or score > best[3]:
            best = (sheet_name, data_df, header_row, score)

    if best is None:
        raise ValueError("Keine passende Tabelle mit Kopfzeile gefunden.")
    return best[0], best[1], best[2]


def load_questions_from_excel(uploaded_file) -> Tuple[Dict[str, Question], str]:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_name, df, _ = detect_best_sheet(excel_file)
    col_map = {normalize_key(c): c for c in df.columns}

    def col(*candidates: str) -> Optional[str]:
        for candidate in candidates:
            key = normalize_key(candidate)
            if key in col_map:
                return col_map[key]
        return None

    question_no_col = col("Frage Nr.")
    question_col = col("Frage")
    block_col = col("Block")
    answer_count_col = col("Anzahl der Antworten")
    input_type_col = col("Antwortart")

    answerart_candidates = [c for c in df.columns if normalize_key(c) == "antwortart"]
    choice_type_col = answerart_candidates[1] if len(answerart_candidates) >= 2 else input_type_col

    if question_no_col is None or question_col is None:
        raise ValueError("Pflichtspalten 'Frage Nr.' oder 'Frage' fehlen.")

    questions: Dict[str, Question] = {}

    for _, row in df.iterrows():
        q_no = normalize_question_no(row.get(question_no_col))
        q_text = normalize_text(row.get(question_col))
        if not q_no or not q_text:
            continue

        answer_count = 0
        if answer_count_col is not None:
            raw_count = row.get(answer_count_col)
            try:
                if raw_count is not None and not pd.isna(raw_count):
                    answer_count = int(float(raw_count))
            except Exception:
                answer_count = 0

        input_type = normalize_text(row.get(input_type_col)) if input_type_col else ""
        choice_type = normalize_text(row.get(choice_type_col)) if choice_type_col else ""

        options: List[AnswerOption] = []

        for answer_col, next_col in zip(ANSWER_TEXT_COLS, ANSWER_NEXT_COLS):
            if answer_col in df.columns:
                answer_text = normalize_text(row.get(answer_col))
                next_question = normalize_question_no(row.get(next_col)) if next_col in df.columns else ""
                if answer_text:
                    options.append(AnswerOption(text=answer_text, next_question=next_question or None))

        if not options:
            for answer_key, next_key in zip(ALT_ANSWER_TEXT_KEYS, ALT_ANSWER_NEXT_KEYS):
                answer_col = col(answer_key)
                next_col = col(next_key)
                if answer_col:
                    answer_text = normalize_text(row.get(answer_col))
                    next_question = normalize_question_no(row.get(next_col)) if next_col else ""
                    if answer_text:
                        options.append(AnswerOption(text=answer_text, next_question=next_question or None))

        if not input_type:
            input_type = "Eingabefeld" if not options else "Button"
        if not choice_type:
            choice_type = "multiple Choice" if "mehrfach" in normalize_key(input_type) else "one Choice"

        questions[q_no] = Question(
            number=q_no,
            block=normalize_text(row.get(block_col)) if block_col else "",
            text=q_text,
            answer_count=answer_count,
            input_type=input_type,
            choice_type=choice_type,
            options=options,
            source_sheet=sheet_name,
        )

    if "1" not in questions:
        first_key = next(iter(questions.keys()), None)
        if first_key:
            questions["1"] = questions[first_key]
        else:
            raise ValueError("Keine Fragen gefunden.")

    return questions, sheet_name


def init_state() -> None:
    defaults = {
        "questions": None,
        "current_question_no": None,
        "history": [],
        "finished": False,
        "uploaded_filename": None,
        "sheet_name": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_questionnaire(keep_questions: bool = False) -> None:
    questions = st.session_state.get("questions") if keep_questions else None
    filename = st.session_state.get("uploaded_filename") if keep_questions else None
    sheet_name = st.session_state.get("sheet_name") if keep_questions else None
    st.session_state["questions"] = questions
    st.session_state["uploaded_filename"] = filename
    st.session_state["sheet_name"] = sheet_name
    st.session_state["current_question_no"] = "1" if questions else None
    st.session_state["history"] = []
    st.session_state["finished"] = False


def determine_next_question(question: Question, answer_value: Any) -> Optional[str]:
    if isinstance(answer_value, list):
        for selected in answer_value:
            for option in question.options:
                if option.text == selected and option.next_question:
                    return option.next_question
        return None

    answer_value = normalize_text(answer_value)
    for option in question.options:
        if option.text == answer_value:
            return option.next_question
    return None


def save_answer(question: Question, answer_value: Any) -> None:
    next_question = determine_next_question(question, answer_value)

    st.session_state.history.append(
        {
            "frage_nr": question.number,
            "block": question.block,
            "frage": question.text,
            "antwort": ", ".join(answer_value) if isinstance(answer_value, list) else normalize_text(answer_value),
            "naechste_frage": next_question or "Ende",
        }
    )

    if next_question and next_question in st.session_state.questions:
        st.session_state.current_question_no = next_question
    else:
        st.session_state.current_question_no = None
        st.session_state.finished = True


def go_back_one_step() -> None:
    if not st.session_state.history:
        return
    st.session_state.history.pop()
    if not st.session_state.history:
        st.session_state.current_question_no = "1"
    else:
        st.session_state.current_question_no = st.session_state.history[-1]["naechste_frage"]
        if st.session_state.current_question_no == "Ende":
            st.session_state.current_question_no = st.session_state.history[-1]["frage_nr"]
    st.session_state.finished = False


def render_question_banner(title: str, question_no: str, block: str = "") -> None:
    block_html = f"<div class='question-meta'>Block: {block}</div>" if block else ""
    st.markdown(
        f"""
        <div class="question-banner">
            <div class="question-label">Frage {title}</div>
            <div class="question-title">{question_no}</div>
            {block_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


st.set_page_config(page_title="Frage-Antwort-Mechanismus", layout="wide")
init_state()

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f7f8fc;
    }
    .block-container {
        max-width: 100% !important;
        padding-top: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    .question-banner {
        width: 100%;
        background: #0b2e59;
        color: white;
        border-radius: 0.9rem;
        padding: 1.1rem 1.4rem;
        margin: 0.4rem 0 1.2rem 0;
        box-shadow: 0 2px 10px rgba(11, 46, 89, 0.18);
    }
    .question-label {
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.3rem;
    }
    .question-title {
        font-size: 1.2rem;
        line-height: 1.5;
        font-weight: 600;
    }
    .question-meta {
        margin-top: 0.45rem;
        font-size: 0.92rem;
        opacity: 0.9;
    }
    div.stButton > button {
        background-color: #ffd8a8;
        color: #000000;
        border: 1px solid #f0b56b;
        border-radius: 0.7rem;
        width: 100%;
        font-weight: 600;
    }
    div.stButton > button:hover {
        background-color: #ffcf93;
        color: #000000;
        border: 1px solid #e8a95a;
    }
    div.stDownloadButton > button {
        background-color: #ffd8a8;
        color: #000000;
        border: 1px solid #f0b56b;
        border-radius: 0.7rem;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Dynamischer Frage-Antwort-Mechanismus")
st.write("Lade eine Excel-Datei hoch. Die App erkennt automatisch das passende Tabellenblatt und verzweigt anhand der Excel-Werte.")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xlsm", "xls"])

if uploaded_file is not None:
    reload_needed = st.session_state.uploaded_filename != uploaded_file.name or st.session_state.questions is None
    if reload_needed:
        try:
            questions, sheet_name = load_questions_from_excel(uploaded_file)
            st.session_state.questions = questions
            st.session_state.uploaded_filename = uploaded_file.name
            st.session_state.sheet_name = sheet_name
            reset_questionnaire(keep_questions=True)
            st.success(f"Datei '{uploaded_file.name}' erfolgreich geladen. {len(questions)} Fragen erkannt. Verwendetes Blatt: {sheet_name}")
        except Exception as exc:
            st.session_state.questions = None
            st.error(f"Fehler beim Einlesen der Excel-Datei: {exc}")

if st.session_state.questions:
    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("Neu starten"):
            reset_questionnaire(keep_questions=True)
            st.rerun()
    with c2:
        if st.session_state.history and st.button("Eine Frage zurück"):
            go_back_one_step()
            st.rerun()

    if not st.session_state.finished and st.session_state.current_question_no:
        question_no = st.session_state.current_question_no
        if question_no not in st.session_state.questions:
            st.error(f"Die nächste Frage '{question_no}' wurde im ausgewählten Tabellenblatt nicht gefunden.")
            st.stop()

        question = st.session_state.questions[question_no]
        question_index = len(st.session_state.history) + 1

        render_question_banner(str(question_index), question.text, question.block)
        st.caption(f"Interne Frage-Nr.: {question.number}")
        st.caption(f"Tabellenblatt: {question.source_sheet}")

        input_type = normalize_key(question.input_type)
        choice_type = normalize_key(question.choice_type)
        option_texts = [option.text for option in question.options]

        if "eingabe" in input_type or "ganze zahl" in input_type or not option_texts:
            is_number = "zahl" in input_type or "anzahl" in normalize_key(question.text) or "wieviel" in normalize_key(question.text)
            if is_number:
                value = st.number_input("Antwort eingeben", step=1, min_value=0, key=f"input_{question.number}")
                if st.button("Weiter", key=f"continue_{question.number}"):
                    save_answer(question, str(int(value)))
                    st.rerun()
            else:
                value = st.text_input("Antwort eingeben", key=f"input_{question.number}")
                if st.button("Weiter", key=f"continue_{question.number}"):
                    if not str(value).strip():
                        st.warning("Bitte erst eine Antwort eingeben.")
                    else:
                        save_answer(question, value)
                        st.rerun()
        elif "multiple" in choice_type or "mehrfach" in choice_type:
            selected = st.multiselect("Eine oder mehrere Antworten auswählen", options=option_texts, key=f"multi_{question.number}")
            if st.button("Weiter", key=f"multi_continue_{question.number}"):
                if not selected:
                    st.warning("Bitte mindestens eine Antwort auswählen.")
                else:
                    save_answer(question, selected)
                    st.rerun()
        elif "button" in input_type:
            st.write("Antwort anklicken:")
            for option in question.options:
                if st.button(option.text, key=f"btn_{question.number}_{option.text}"):
                    save_answer(question, option.text)
                    st.rerun()
        else:
            selected = st.radio("Antwort auswählen", options=option_texts, key=f"single_{question.number}")
            if st.button("Weiter", key=f"radio_continue_{question.number}"):
                save_answer(question, selected)
                st.rerun()

        if option_texts:
            st.caption(f"Antwortoptionen: {len(option_texts)}")

    if st.session_state.history:
        st.divider()
        st.subheader("Bisherige Antworten")
        history_df = pd.DataFrame(st.session_state.history)
        st.dataframe(history_df, use_container_width=True)
        st.download_button(
            "Antworten als CSV herunterladen",
            data=history_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="antworten.csv",
            mime="text/csv",
        )

    if st.session_state.finished:
        st.success("Fragebogen beendet.")
        st.info("Hinweis: In der Excel-Datei enden einige Pfade ohne definierte Folgennummer. Dann beendet die App den Fragebogen automatisch.")
