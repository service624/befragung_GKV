from __future__ import annotations

import re
import json
from pathlib import Path
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

ANSWER_TEXT_COLS = ["I", "K", "M", "O", "Q", "S"]
ANSWER_NEXT_COLS = ["J", "L", "N", "P", "R", "T"]
ALT_ANSWER_TEXT_KEYS = ["a1", "a2", "a3", "a4", "a5", "a6"]
ALT_ANSWER_NEXT_KEYS = ["nach a1", "nach a2", "nach a3", "nach a4", "nach a5", "nach a6"]


SETTINGS_FILE = Path(__file__).with_name("GKV_settings.json")
SAVED_UPLOAD_DIR = Path(__file__).with_name("gkv_saved_uploads")


def load_app_settings() -> Dict[str, Any]:
    defaults: Dict[str, Any] = {
        "app_title": "Dynamischer Frage-Antwort-Mechanismus",
        "app_subtitle": "Excel-Datei hochladen, Fragenpfad automatisch durchlaufen und Antworten am Ende als CSV oder Excel exportieren.",
        "show_header": True,
        "show_subheader": True,
        "upload_label": "Excel-Datei hochladen",
        "saved_upload_path": "",
        "saved_upload_name": "",
        "expander_closed_bg": "transparent",
        "expander_closed_text": "#0a2b55",
        "expander_open_bg": "#0a2b55",
        "expander_open_text": "#ffffff",
    }
    try:
        if SETTINGS_FILE.exists():
            raw = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                defaults.update(raw)
    except Exception:
        pass
    return defaults


def save_app_settings(settings: Dict[str, Any]) -> None:
    SETTINGS_FILE.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def persist_uploaded_file(uploaded_file, settings: Dict[str, Any]) -> Dict[str, Any]:
    SAVED_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    suffix = Path(str(getattr(uploaded_file, "name", "") or "upload.xlsx")).suffix or ".xlsx"
    target = SAVED_UPLOAD_DIR / f"last_upload{suffix}"
    with open(target, "wb") as f:
        f.write(uploaded_file.getbuffer())

    updated = dict(settings)
    updated["saved_upload_path"] = str(target)
    updated["saved_upload_name"] = str(getattr(uploaded_file, "name", "") or target.name)
    save_app_settings(updated)
    return updated


def get_saved_upload_path(settings: Dict[str, Any]) -> Optional[Path]:
    raw = str(settings.get("saved_upload_path", "") or "").strip()
    if not raw:
        return None
    path = Path(raw)
    return path if path.exists() and path.is_file() else None


def clear_saved_upload(settings: Dict[str, Any]) -> Dict[str, Any]:
    updated = dict(settings)
    saved_path = get_saved_upload_path(updated)
    if saved_path and saved_path.exists():
        try:
            saved_path.unlink()
        except Exception:
            pass
    updated["saved_upload_path"] = ""
    updated["saved_upload_name"] = ""
    save_app_settings(updated)
    return updated


def render_settings_expander(settings: Dict[str, Any]) -> Tuple[Optional[Any], Dict[str, Any]]:
    uploaded_source = None
    draft = dict(settings)

    with st.expander("Einstellungen", expanded=False):
        tabs = st.tabs(["Upload", "Header"])

        with tabs[0]:
            st.markdown("### Upload")
            uploaded_file = st.file_uploader(
                str(settings.get("upload_label", "Excel-Datei hochladen")),
                type=["xlsx", "xlsm", "xls"],
                key="settings_embedded_uploader",
            )

            if uploaded_file is not None:
                settings = persist_uploaded_file(uploaded_file, settings)
                draft = dict(settings)
                st.success(f"Datei gespeichert: {settings.get('saved_upload_name', '')}")
                uploaded_source = uploaded_file

            saved_path = get_saved_upload_path(settings)
            saved_name = str(settings.get("saved_upload_name", "") or "")
            info_col, action_col = st.columns([4, 1.4])
            if saved_path:
                info_col.caption(f"Gespeicherte Datei: {saved_name or saved_path.name}")
                if action_col.button("Gespeicherte Datei entfernen", key="clear_saved_upload_btn", use_container_width=True):
                    settings = clear_saved_upload(settings)
                    draft = dict(settings)
                    st.success("Gespeicherte Datei entfernt.")
                    st.rerun()
                if uploaded_source is None:
                    uploaded_source = str(saved_path)
            else:
                info_col.caption("Keine Datei gespeichert.")

        with tabs[1]:
            st.markdown("### Header")
            draft["app_title"] = st.text_input(
                "Headername",
                value=str(settings.get("app_title", "")),
                key="cfg_app_title",
            )
            draft["app_subtitle"] = st.text_area(
                "Subheader",
                value=str(settings.get("app_subtitle", "")),
                key="cfg_app_subtitle",
                height=100,
            )
            c1, c2 = st.columns(2)
            draft["show_header"] = c1.checkbox(
                "Header anzeigen",
                value=bool(settings.get("show_header", True)),
                key="cfg_show_header",
            )
            draft["show_subheader"] = c2.checkbox(
                "Subheader anzeigen",
                value=bool(settings.get("show_subheader", True)),
                key="cfg_show_subheader",
            )
            st.caption("Expander-Farbe: zugeklappt transparent, aufgeklappt dunkelblau.")

        save_col, _ = st.columns([1.2, 5])
        if save_col.button("Einstellungen speichern", key="save_settings_btn", use_container_width=True):
            save_app_settings(draft)
            settings = dict(draft)
            st.success("Einstellungen gespeichert.")
            st.rerun()

    return uploaded_source, settings


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
    default_next_question: Optional[str] = None
    input_format: str = ""


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


def parse_numeric_input(value: Any, allow_float: bool = True) -> Optional[float]:
    text = normalize_text(value)
    if not text:
        return None

    cleaned = text.replace("€", "").replace("EUR", "").replace("%", "")
    cleaned = cleaned.replace(" ", "")
    cleaned = cleaned.replace(".", "").replace(",", ".")
    cleaned = cleaned.strip()

    try:
        number = float(cleaned)
    except Exception:
        return None

    if not allow_float:
        if not number.is_integer():
            return None
        return int(number)
    return number


def find_header_row(df: pd.DataFrame) -> Optional[int]:
    for idx in range(min(len(df), 30)):
        row_values = [normalize_key(v) for v in df.iloc[idx].tolist()]
        if "frage nr." in row_values and "frage" in row_values:
            return idx
    return None


def detect_best_sheet(excel_file: pd.ExcelFile) -> Tuple[str, pd.DataFrame]:
    best: Optional[Tuple[str, pd.DataFrame, int]] = None

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

        if best is None or score > best[2]:
            best = (sheet_name, data_df, score)

    if best is None:
        raise ValueError("Keine passende Tabelle mit Kopfzeile gefunden.")
    return best[0], best[1]


def load_questions_from_excel(uploaded_file) -> Tuple[Dict[str, Question], str]:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_name, df = detect_best_sheet(excel_file)
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

    answerart_candidates = [c for c in df.columns if normalize_key(c) == "antwortart"]
    input_type_col = answerart_candidates[0] if answerart_candidates else None
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
        input_format = choice_type if normalize_key(input_type) == "eingabefeld" else ""

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

        default_next_question = None
        if not options:
            first_next_col = next((col_name for col_name in ANSWER_NEXT_COLS if col_name in df.columns), None)

            if first_next_col is None:
                for next_key in ALT_ANSWER_NEXT_KEYS:
                    mapped_next_col = col(next_key)
                    if mapped_next_col:
                        first_next_col = mapped_next_col
                        break

            if first_next_col:
                default_next_question = normalize_question_no(row.get(first_next_col)) or None

        if not input_type:
            input_type = "Eingabefeld" if not options else "Button"
        if not choice_type:
            choice_type = "multiple Choice" if "mehrfach" in normalize_key(input_type) else "one Choice"
        if normalize_key(input_type) != "eingabefeld":
            input_format = ""

        questions[q_no] = Question(
            number=q_no,
            block=normalize_text(row.get(block_col)) if block_col else "",
            text=q_text,
            answer_count=answer_count,
            input_type=input_type,
            choice_type=choice_type,
            options=options,
            source_sheet=sheet_name,
            default_next_question=default_next_question,
            input_format=input_format,
        )

    if not questions:
        raise ValueError("Keine Fragen gefunden.")
    if "1" not in questions:
        first_key = next(iter(questions.keys()))
        questions["1"] = questions[first_key]

    return questions, sheet_name


def init_state() -> None:
    defaults = {
        "questions": None,
        "current_question_no": None,
        "history": [],
        "answers": {},
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
    st.session_state["answers"] = {}
    st.session_state["finished"] = False

    for key in list(st.session_state.keys()):
        if key.startswith(("input_", "single_", "multi_", "cb_")):
            del st.session_state[key]


def clear_widget_state_for_question(question_no: str) -> None:
    prefixes = [f"input_{question_no}", f"single_{question_no}", f"multi_{question_no}", f"cb_{question_no}_"]
    for key in list(st.session_state.keys()):
        if any(key.startswith(prefix) for prefix in prefixes):
            del st.session_state[key]


def determine_next_question(question: Question, answer_value: Any) -> Optional[str]:
    if isinstance(answer_value, list):
        for selected in answer_value:
            for option in question.options:
                if option.text == selected and option.next_question:
                    return option.next_question
        return question.default_next_question

    value = normalize_text(answer_value)
    for option in question.options:
        if option.text == value:
            return option.next_question

    return question.default_next_question


def save_answer(question: Question, answer_value: Any) -> None:
    next_question = determine_next_question(question, answer_value)
    answer_text = ", ".join(answer_value) if isinstance(answer_value, list) else normalize_text(answer_value)

    st.session_state.answers[question.number] = answer_value
    st.session_state.history.append(
        {
            "frage_nr": question.number,
            "block": question.block,
            "frage": question.text,
            "antwort": answer_text,
            "naechste_frage": next_question or "Ende",
        }
    )

    if next_question and next_question in st.session_state.questions:
        st.session_state.current_question_no = next_question
        st.session_state.finished = False
    else:
        st.session_state.current_question_no = None
        st.session_state.finished = True


def go_back_one_step() -> None:
    if not st.session_state.history:
        st.session_state.current_question_no = "1"
        st.session_state.finished = False
        return

    last_entry = st.session_state.history.pop()
    question_no = last_entry["frage_nr"]
    st.session_state.answers.pop(question_no, None)
    clear_widget_state_for_question(question_no)
    st.session_state.current_question_no = question_no
    st.session_state.finished = False


def reachable_question_count(questions: Dict[str, Question], start: str = "1") -> int:
    visited: set[str] = set()
    stack = [start]
    while stack:
        q_no = stack.pop()
        if q_no in visited or q_no not in questions:
            continue
        visited.add(q_no)

        current = questions[q_no]
        for option in current.options:
            if option.next_question:
                stack.append(option.next_question)

        if current.default_next_question:
            stack.append(current.default_next_question)

    return len(visited) if visited else len(questions)


def render_question_banner(step_label: str, question_text: str, block: str = "") -> None:
    block_html = f"<div class='question-meta'>Block: {block}</div>" if block else ""
    st.markdown(
        f"""
        <div class="question-banner">
            <div class="question-label">{step_label}</div>
            <div class="question-title">{question_text}</div>
            {block_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_progress() -> None:
    questions = st.session_state.questions or {}
    total = reachable_question_count(questions)
    answered = len(st.session_state.history)
    current_index = answered + (0 if st.session_state.finished else 1)
    progress_value = 1.0 if st.session_state.finished else min(answered / max(total, 1), 1.0)

    st.progress(progress_value)
    if st.session_state.finished:
        st.caption(f"Abgeschlossen: {answered} von {total} erreichbaren Fragen beantwortet")
    else:
        st.caption(f"Fortschritt: Frage {current_index} von ca. {total}")


def answers_to_excel_bytes(history_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        history_df.to_excel(writer, sheet_name="Antworten", index=False)
    output.seek(0)
    return output.getvalue()


def render_button_options(question: Question) -> None:
    options = question.options
    per_row = min(4, max(1, len(options)))
    for start in range(0, len(options), per_row):
        row_options = options[start : start + per_row]
        cols = st.columns(per_row)
        for idx, option in enumerate(row_options):
            absolute_idx = start + idx
            with cols[idx]:
                if st.button(
                    option.text,
                    key=f"btn_{question.number}_{absolute_idx}",
                    use_container_width=True,
                ):
                    save_answer(question, option.text)
                    st.rerun()


def render_single_choice(question: Question) -> None:
    options = [opt.text for opt in question.options]
    selected = st.radio(
        "Antwort auswählen",
        options=options,
        key=f"single_{question.number}",
        horizontal=True,
        label_visibility="collapsed",
    )
    if st.button("Weiter", key=f"single_continue_{question.number}"):
        save_answer(question, selected)
        st.rerun()


def render_multiple_choice(question: Question) -> None:
    options = [opt.text for opt in question.options]
    per_row = min(4, max(1, len(options)))

    for start in range(0, len(options), per_row):
        row_options = options[start : start + per_row]
        cols = st.columns(per_row)
        for idx, option in enumerate(row_options):
            with cols[idx]:
                st.checkbox(option, key=f"cb_{question.number}_{start + idx}")

    if st.button("Weiter", key=f"multi_continue_{question.number}"):
        selected = [
            option
            for idx, option in enumerate(options)
            if st.session_state.get(f"cb_{question.number}_{idx}", False)
        ]
        if not selected:
            st.warning("Bitte mindestens eine Antwort auswählen.")
        else:
            save_answer(question, selected)
            st.rerun()


def render_input(question: Question) -> None:
    input_kind = normalize_key(question.input_format or question.choice_type or question.input_type)
    text_key = f"input_{question.number}"
    num_key = f"input_num_{question.number}"

    if "ganze zahl" in input_kind:
        value = st.number_input("Antwort eingeben", step=1, format="%d", value=0, key=num_key)
        if st.button("Weiter", key=f"continue_{question.number}"):
            save_answer(question, str(int(value)))
            st.rerun()
        return

    if input_kind in {"zahl", "eur", "%", "prozent"} or "zahl" in input_kind or "eur" in input_kind or "%" in input_kind:
        label = "Antwort eingeben"
        if "eur" in input_kind:
            label = "Antwort eingeben (EUR)"
        elif "%" in input_kind or "prozent" in input_kind:
            label = "Antwort eingeben (%)"

        value = st.text_input(label, key=text_key, value="0", placeholder="z. B. 1234,56")
        if st.button("Weiter", key=f"continue_{question.number}"):
            parsed = parse_numeric_input(value, allow_float=True)
            if parsed is None:
                st.warning("Bitte eine gültige Zahl eingeben.")
            else:
                if "ganze" in input_kind or parsed.is_integer():
                    out = str(int(parsed))
                else:
                    out = str(parsed).replace(".", ",")
                save_answer(question, out)
                st.rerun()
        return

    value = st.text_input("Antwort eingeben", key=text_key)
    if st.button("Weiter", key=f"continue_{question.number}"):
        if not str(value).strip():
            st.warning("Bitte erst eine Antwort eingeben.")
        else:
            save_answer(question, value)
            st.rerun()


def render_result_page() -> None:
    history = st.session_state.history
    history_df = pd.DataFrame(history)
    blocks = history_df["block"].replace("", pd.NA).dropna().nunique() if not history_df.empty else 0
    last_question = history[-1]["frage"] if history else "-"
    last_answer = history[-1]["antwort"] if history else "-"

    st.success("Fragebogen beendet")
    c1, c2, c3 = st.columns(3)
    c1.metric("Beantwortete Fragen", len(history))
    c2.metric("Blöcke durchlaufen", int(blocks))
    c3.metric("Letzte interne Frage-Nr.", history[-1]["frage_nr"] if history else "-")

    st.markdown("### Ergebnisübersicht")
    st.info(f"Letzte Frage: {last_question}\n\nLetzte Antwort: {last_answer}")

    with st.expander("Antwortverlauf anzeigen", expanded=True):
        if not history_df.empty:
            st.dataframe(history_df, use_container_width=True, hide_index=True)
        else:
            st.write("Es wurden noch keine Antworten gespeichert.")

    if not history_df.empty:
        csv_bytes = history_df.to_csv(index=False).encode("utf-8-sig")
        xlsx_bytes = answers_to_excel_bytes(history_df)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Antworten als CSV herunterladen",
                data=csv_bytes,
                file_name="antworten.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Antworten als Excel herunterladen",
                data=xlsx_bytes,
                file_name="antworten.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.caption("Hinweis: Wenn in der Excel-Datei für eine Antwort keine nächste Frage hinterlegt ist, endet der Pfad automatisch.")


st.set_page_config(page_title="Frage-Antwort-Mechanismus", layout="wide")
init_state()
app_settings = load_app_settings()

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f6f8fc;
    }
    .block-container {
        max-width: 100% !important;
        padding-top: 2.2rem;
        padding-left: 1.6rem;
        padding-right: 1.6rem;
        padding-bottom: 2rem;
    }
    details > summary {
        background: transparent !important;
        color: #0a2b55 !important;
        border-radius: 0.8rem !important;
        border: 1px solid rgba(10, 43, 85, 0.14) !important;
        padding: 0.55rem 0.85rem !important;
    }
    details[open] > summary {
        background: #0a2b55 !important;
        color: #ffffff !important;
    }
    .question-banner {
        width: 100%;
        background: linear-gradient(90deg, #0a2b55 0%, #123f78 100%);
        color: white;
        border-radius: 0.9rem;
        padding: 1rem 1.35rem;
        margin: 0.25rem 0 1rem 0;
        box-shadow: 0 4px 14px rgba(10, 43, 85, 0.16);
    }
    .question-label {
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
        opacity: 0.96;
    }
    .question-title {
        font-size: 1.28rem;
        line-height: 1.45;
        font-weight: 600;
    }
    .question-meta {
        margin-top: 0.5rem;
        font-size: 0.95rem;
        opacity: 0.92;
    }
    div.stButton > button,
    div.stDownloadButton > button {
        background-color: #ffd6ad !important;
        color: #000000 !important;
        border: 1px solid #efb167 !important;
        border-radius: 0.8rem !important;
        width: 100% !important;
        font-weight: 600 !important;
        min-height: 2.9rem !important;
    }
    div.stButton > button:hover,
    div.stDownloadButton > button:hover {
        background-color: #ffcc94 !important;
        color: #000000 !important;
        border: 1px solid #e1a255 !important;
    }
    div[data-baseweb="radio"] > div {
        gap: 0.65rem;
        flex-wrap: wrap;
    }
    div[data-baseweb="radio"] label {
        background: #fff3e5;
        border: 1px solid #f1c48e;
        border-radius: 0.8rem;
        padding: 0.45rem 0.85rem;
    }
    div[data-testid="stCheckbox"] {
        background: #fff3e5;
        border: 1px solid #f1c48e;
        border-radius: 0.8rem;
        padding: 0.35rem 0.7rem;
        margin-bottom: 0.55rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if app_settings.get("show_header", True):
    st.title(str(app_settings.get("app_title", "Dynamischer Frage-Antwort-Mechanismus")))
if app_settings.get("show_subheader", True):
    st.write(str(app_settings.get("app_subtitle", "Excel-Datei hochladen, Fragenpfad automatisch durchlaufen und Antworten am Ende als CSV oder Excel exportieren.")))

uploaded_file, app_settings = render_settings_expander(app_settings)

if uploaded_file is not None:
    current_filename = uploaded_file.name if hasattr(uploaded_file, "name") else Path(str(uploaded_file)).name
    reload_needed = st.session_state.uploaded_filename != current_filename or st.session_state.questions is None
    if reload_needed:
        try:
            questions, sheet_name = load_questions_from_excel(uploaded_file)
            st.session_state.questions = questions
            st.session_state.uploaded_filename = current_filename
            st.session_state.sheet_name = sheet_name
            reset_questionnaire(keep_questions=True)
            st.success(
                f"Datei '{current_filename}' geladen. {len(questions)} Fragen erkannt. Verwendetes Blatt: {sheet_name}"
            )
        except Exception as exc:
            st.session_state.questions = None
            st.error(f"Fehler beim Einlesen der Excel-Datei: {exc}")

if st.session_state.questions:
    render_progress()

    top_left, top_mid, top_right = st.columns([1, 1, 2])
    with top_left:
        if st.button("Neu starten", use_container_width=True):
            reset_questionnaire(keep_questions=True)
            st.rerun()
    with top_mid:
        back_disabled = not st.session_state.history
        if st.button("Eine Frage zurück", disabled=back_disabled, use_container_width=True):
            go_back_one_step()
            st.rerun()
    with top_right:
        st.caption(f"Datei: {st.session_state.uploaded_filename} | Tabellenblatt: {st.session_state.sheet_name}")

    if not st.session_state.finished and st.session_state.current_question_no:
        question_no = st.session_state.current_question_no
        questions = st.session_state.questions
        if question_no not in questions:
            st.error(f"Die nächste Frage '{question_no}' wurde im ausgewählten Tabellenblatt nicht gefunden.")
            st.stop()

        question = questions[question_no]
        question_index = len(st.session_state.history) + 1
        render_question_banner(f"Frage {question_index}", question.text, question.block)
        st.caption(f"Interne Frage-Nr.: {question.number}")

        input_type = normalize_key(question.input_type)
        choice_type = normalize_key(question.choice_type)
        option_texts = [option.text for option in question.options]

        if "eingabe" in input_type or "ganze zahl" in input_type or not option_texts:
            render_input(question)
        elif "multiple" in choice_type or "mehrfach" in choice_type:
            render_multiple_choice(question)
        elif "button" in input_type:
            render_button_options(question)
        else:
            render_single_choice(question)

        if option_texts:
            st.caption(f"Antwortoptionen: {len(option_texts)}")

    if st.session_state.finished:
        render_result_page()
    elif st.session_state.history:
        with st.expander("Bisherige Antworten", expanded=False):
            history_df = pd.DataFrame(st.session_state.history)
            st.dataframe(history_df, use_container_width=True, hide_index=True)
