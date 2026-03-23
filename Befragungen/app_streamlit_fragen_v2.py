from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st

ANSWER_TEXT_COLS = ["I", "K", "M", "O", "Q", "S"]
ANSWER_NEXT_COLS = ["J", "L", "N", "P", "R", "T"]


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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def normalize_key(value: Any) -> str:
    return normalize_text(value).lower()


def normalize_question_no(value: Any) -> str:
    return normalize_text(value)


def find_header_row(df: pd.DataFrame) -> int:
    for idx in range(min(len(df), 30)):
        row_values = [normalize_key(v) for v in df.iloc[idx].tolist()]
        if "frage nr." in row_values and "frage" in row_values:
            return idx
    raise ValueError("Keine Kopfzeile mit 'Frage Nr.' und 'Frage' gefunden.")


def load_questions_from_excel(uploaded_file) -> Dict[str, Question]:
    raw_df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    header_row = find_header_row(raw_df)
    headers = [normalize_text(v) or f"col_{i}" for i, v in enumerate(raw_df.iloc[header_row].tolist())]
    df = raw_df.iloc[header_row + 1 :].copy().reset_index(drop=True)
    df.columns = headers

    col_map = {normalize_key(c): c for c in df.columns}

    def col(*candidates: str) -> str:
        for candidate in candidates:
            key = normalize_key(candidate)
            if key in col_map:
                return col_map[key]
        raise KeyError(f"Spalte nicht gefunden. Erwartet eine von: {candidates}")

    question_no_col = col("Frage Nr.")
    block_col = col("Block")
    question_col = col("Frage")
    answer_count_col = col("Anzahl der Antworten")
    input_type_col = col("Antwortart")

    answerart_candidates = [c for c in df.columns if normalize_key(c) == "antwortart"]
    if len(answerart_candidates) >= 2:
        choice_type_col = answerart_candidates[1]
    else:
        choice_type_col = input_type_col

    questions: Dict[str, Question] = {}

    for _, row in df.iterrows():
        q_no = normalize_question_no(row.get(question_no_col))
        q_text = normalize_text(row.get(question_col))
        if not q_no or not q_text:
            continue

        answer_count_raw = row.get(answer_count_col)
        try:
            answer_count = int(float(answer_count_raw)) if pd.notna(answer_count_raw) else 0
        except Exception:
            answer_count = 0

        input_type = normalize_text(row.get(input_type_col)) or "Button"
        choice_type = normalize_text(row.get(choice_type_col)) or "one Choice"

        options: List[AnswerOption] = []
        for answer_col, next_col in zip(ANSWER_TEXT_COLS, ANSWER_NEXT_COLS):
            answer_text = normalize_text(row.get(answer_col))
            next_question = normalize_question_no(row.get(next_col))
            if answer_text:
                options.append(AnswerOption(text=answer_text, next_question=next_question or None))

        questions[q_no] = Question(
            number=q_no,
            block=normalize_text(row.get(block_col)),
            text=q_text,
            answer_count=answer_count,
            input_type=input_type,
            choice_type=choice_type,
            options=options,
        )

    if "1" not in questions:
        raise ValueError("Die erste Frage mit Nummer 1 wurde nicht gefunden.")
    return questions


def init_state() -> None:
    defaults = {
        "questions": None,
        "current_question_no": None,
        "history": [],
        "finished": False,
        "uploaded_filename": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_questionnaire(keep_questions: bool = False) -> None:
    questions = st.session_state.get("questions") if keep_questions else None
    filename = st.session_state.get("uploaded_filename") if keep_questions else None
    st.session_state["questions"] = questions
    st.session_state["uploaded_filename"] = filename
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
    if st.session_state.history:
        st.session_state.current_question_no = st.session_state.history[-1]["naechste_frage"]
        if st.session_state.current_question_no == "Ende":
            st.session_state.current_question_no = st.session_state.history[-1]["frage_nr"]
    else:
        st.session_state.current_question_no = "1"
    st.session_state.finished = False


st.set_page_config(page_title="Frage-Antwort-Mechanismus", layout="centered")
init_state()

st.title("Dynamischer Frage-Antwort-Mechanismus")
st.write("Lade eine Excel-Datei hoch. Die Fragen werden aus Tabelle 1 gelesen und abhängig von der Antwort verzweigt.")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xlsm", "xls"])

if uploaded_file is not None:
    reload_needed = st.session_state.uploaded_filename != uploaded_file.name or st.session_state.questions is None
    if reload_needed:
        try:
            questions = load_questions_from_excel(uploaded_file)
            st.session_state.questions = questions
            st.session_state.uploaded_filename = uploaded_file.name
            reset_questionnaire(keep_questions=True)
            st.success(f"Datei '{uploaded_file.name}' erfolgreich geladen. {len(questions)} Fragen erkannt.")
        except Exception as exc:
            st.session_state.questions = None
            st.error(f"Fehler beim Einlesen der Excel-Datei: {exc}")

if st.session_state.questions:
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("Neu starten"):
            reset_questionnaire(keep_questions=True)
            st.rerun()
    with col2:
        if st.session_state.history and st.button("Eine Frage zurück"):
            go_back_one_step()
            st.rerun()

    if not st.session_state.finished and st.session_state.current_question_no:
        question_no = st.session_state.current_question_no
        if question_no not in st.session_state.questions:
            st.error(f"Die nächste Frage '{question_no}' wurde in der Excel-Datei nicht gefunden.")
            st.stop()

        question = st.session_state.questions[question_no]
        step_no = len(st.session_state.history) + 1

        st.subheader(f"Schritt {step_no}")
        st.caption(f"Interne Frage-Nr.: {question.number}")
        if question.block:
            st.caption(f"Block: {question.block}")
        st.markdown(f"**{question.text}**")

        input_type = normalize_key(question.input_type)
        choice_type = normalize_key(question.choice_type)
        option_texts = [option.text for option in question.options]

        if "eingabe" in input_type:
            text_value = st.text_input("Antwort eingeben", key=f"input_{question.number}")
            if st.button("Weiter", key=f"continue_{question.number}"):
                if not text_value.strip():
                    st.warning("Bitte erst eine Antwort eingeben.")
                else:
                    save_answer(question, text_value)
                    st.rerun()
        else:
            if not option_texts:
                st.warning("Für diese Frage wurden keine Antwortoptionen gefunden.")
            elif "multiple" in choice_type:
                selected = st.multiselect(
                    "Eine oder mehrere Antworten auswählen",
                    options=option_texts,
                    key=f"multi_{question.number}",
                )
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
                selected = st.radio(
                    "Antwort auswählen",
                    options=option_texts,
                    key=f"single_{question.number}",
                )
                if st.button("Weiter", key=f"radio_continue_{question.number}"):
                    save_answer(question, selected)
                    st.rerun()

    if st.session_state.history:
        st.divider()
        st.subheader("Bisherige Antworten")
        history_df = pd.DataFrame(st.session_state.history)
        st.dataframe(history_df, use_container_width=True)

        csv_data = history_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Antworten als CSV herunterladen",
            data=csv_data,
            file_name="antworten.csv",
            mime="text/csv",
        )

    if st.session_state.finished:
        st.success("Fragebogen beendet.")
