from __future__ import annotations

import re
import json
import hashlib
import ipaddress
from pathlib import Path
from dataclasses import dataclass, field
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

ANSWER_TEXT_COLS = ["I", "K", "M", "O", "Q", "S"]
ANSWER_NEXT_COLS = ["J", "L", "N", "P", "R", "T"]
ALT_ANSWER_TEXT_KEYS = ["a1", "a2", "a3", "a4", "a5", "a6"]
ALT_ANSWER_NEXT_KEYS = ["nach a1", "nach a2", "nach a3", "nach a4", "nach a5", "nach a6"]


SETTINGS_FILE = Path(__file__).with_name("GKV_settings.json")


def sha256_text(text: str) -> str:
    return hashlib.sha256((text or "").encode("utf-8")).hexdigest()


def load_app_settings() -> Dict[str, Any]:
    defaults: Dict[str, Any] = {
        "app_title": "Dynamischer Frage-Antwort-Mechanismus",
        "app_subtitle": "Excel-Datei hochladen, Fragenpfad automatisch durchlaufen und Antworten am Ende als CSV oder Excel exportieren.",
        "show_header": True,
        "show_subheader": True,
        "upload_label": "Excel-Datei hochladen",
        "expander_closed_bg": "#e8eef8",
        "expander_closed_text": "#0a2b55",
        "expander_open_bg": "#0a2b55",
        "expander_open_text": "#ffffff",
        "settings_protection_enabled": False,
        "settings_password_hash": "",
        "settings_hint": "",
        "shortcut_combo": "",
        "shortcut_enabled": False,
        "saved_upload_path": "",
        "saved_upload_name": "",
        "settings_ip_protection_enabled": False,
        "settings_allowed_ips": [],
        "settings_visibility_mode": "ip_or_shortcut",
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


def normalize_ip_value(value: str) -> str:
    return str(value or "").strip().lower()


def parse_ip_lines(text_value: str) -> List[str]:
    return [
        normalize_ip_value(x)
        for x in re.split(r"[\n,;]+", str(text_value or ""))
        if normalize_ip_value(x)
    ]


def is_valid_ip_rule(value: str) -> bool:
    rule = normalize_ip_value(value)
    if not rule:
        return False
    if "*" in rule:
        parts = rule.split(".")
        if len(parts) != 4:
            return False
        for part in parts:
            if part == "*":
                continue
            if not part.isdigit():
                return False
            if not 0 <= int(part) <= 255:
                return False
        return True
    try:
        if "/" in rule:
            ipaddress.ip_network(rule, strict=False)
        else:
            ipaddress.ip_address(rule)
        return True
    except Exception:
        return False


def ip_rule_matches(rule: str, current_ip: str) -> bool:
    rule = normalize_ip_value(rule)
    current_ip = normalize_ip_value(current_ip)
    if not rule or not current_ip:
        return False
    try:
        current_obj = ipaddress.ip_address(current_ip)
    except Exception:
        return False
    if "*" in rule:
        rule_parts = rule.split(".")
        ip_parts = current_ip.split(".")
        if len(rule_parts) != 4 or len(ip_parts) != 4:
            return False
        return all(r == "*" or r == ip for r, ip in zip(rule_parts, ip_parts))
    try:
        if "/" in rule:
            return current_obj in ipaddress.ip_network(rule, strict=False)
        return current_obj == ipaddress.ip_address(rule)
    except Exception:
        return False


def get_current_request_ip() -> str:
    manual_ip = normalize_ip_value(st.session_state.get("manual_client_ip", ""))
    if manual_ip:
        return manual_ip

    headers = None
    try:
        headers = getattr(getattr(st, "context", None), "headers", None)
    except Exception:
        headers = None

    if headers:
        try:
            for key, value in headers.items():
                key_l = str(key).lower()
                if key_l in {"x-forwarded-for", "x-real-ip", "cf-connecting-ip", "forwarded", "remote_addr"}:
                    first = str(value).split(",", 1)[0].strip()
                    if first:
                        return normalize_ip_value(first)
        except Exception:
            pass

    return ""


def settings_visible_for_ip(settings: Dict[str, Any], current_ip: str) -> bool:
    if not bool(settings.get("settings_ip_protection_enabled", False)):
        return True

    allowed_rules = [
        rule for rule in settings.get("settings_allowed_ips", [])
        if is_valid_ip_rule(str(rule))
    ]
    if not allowed_rules:
        return False

    current_ip = normalize_ip_value(current_ip)
    return any(ip_rule_matches(rule, current_ip) for rule in allowed_rules)


def parse_shortcut(combo: str) -> Dict[str, Any]:
    parts = [p.strip().lower() for p in str(combo or "").split("+") if p.strip()]
    main_key = parts[-1] if parts else "s"
    return {
        "ctrl": "ctrl" in parts or "control" in parts,
        "alt": "alt" in parts,
        "shift": "shift" in parts,
        "meta": "meta" in parts or "cmd" in parts or "command" in parts,
        "key": main_key,
    }


def register_settings_shortcut(settings: Dict[str, Any]) -> None:
    if not bool(settings.get("shortcut_enabled", False)):
        return

    combo = parse_shortcut(str(settings.get("shortcut_combo", "") or ""))
    js = f"""
    <script>
    (function() {{
      const cfg = {json.dumps(combo)};
      const storageKey = "gkv_settings_shortcut_trigger";

      if (window.__gkvSettingsShortcutRegistered) return;
      window.__gkvSettingsShortcutRegistered = true;

      function keyMatches(event) {{
        const key = String(event.key || "").toLowerCase();
        return (
          (!!cfg.ctrl === !!event.ctrlKey) &&
          (!!cfg.alt === !!event.altKey) &&
          (!!cfg.shift === !!event.shiftKey) &&
          (!!cfg.meta === !!event.metaKey) &&
          key === String(cfg.key || "").toLowerCase()
        );
      }}

      window.addEventListener("keydown", function(event) {{
        if (!keyMatches(event)) return;
        event.preventDefault();
        try {{
          const next = String(Date.now());
          window.localStorage.setItem(storageKey, next);
        }} catch (e) {{}}
        window.parent.location.reload();
      }});
    }})();
    </script>
    """
    components.html(js, height=0)


def sync_shortcut_unlock() -> None:
    storage_key = "gkv_settings_shortcut_trigger"
    bridge_key = "__gkv_settings_shortcut_value"

    st.text_input(bridge_key, key=bridge_key, label_visibility="collapsed")
    st.markdown(
        f"""
        <style>
        input[aria-label="{bridge_key}"] {{
            display:none !important;
            height:0 !important;
            padding:0 !important;
            border:0 !important;
            margin:0 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    html = f"""
    <script>
    (function() {{
      try {{
        const value = window.localStorage.getItem("{storage_key}") || "";
        const parentDoc = window.parent.document;
        const input = parentDoc.querySelector('input[aria-label="{bridge_key}"]');
        if (!input) return;
        const setter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, 'value')?.set;
        if (setter) setter.call(input, value);
        else input.value = value;
        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }} catch (e) {{}}
    }})();
    </script>
    """
    components.html(html, height=0)

    current_value = str(st.session_state.get(bridge_key, "") or "")
    last_value = str(st.session_state.get("__last_shortcut_unlock_value", "") or "")
    if current_value and current_value != last_value:
        st.session_state["__last_shortcut_unlock_value"] = current_value
        st.session_state["settings_shortcut_override"] = True


def settings_visible(settings: Dict[str, Any], current_ip: str) -> bool:
    visibility_mode = str(settings.get("settings_visibility_mode", "ip_or_shortcut") or "ip_or_shortcut")
    ip_ok = settings_visible_for_ip(settings, current_ip)
    shortcut_ok = bool(st.session_state.get("settings_shortcut_override", False))

    if visibility_mode == "ip_only":
        return ip_ok
    return ip_ok or shortcut_ok


def get_saved_upload_dir() -> Path:
    target = Path(__file__).with_name("gkv_saved_uploads")
    target.mkdir(parents=True, exist_ok=True)
    return target


def get_saved_upload_path(settings: Dict[str, Any]) -> Optional[Path]:
    raw = str((settings or {}).get("saved_upload_path", "") or "").strip()
    if not raw:
        return None
    path = Path(raw)
    return path if path.exists() and path.is_file() else None


def persist_uploaded_file(uploaded_file: Any, settings: Dict[str, Any]) -> Dict[str, Any]:
    suffix = Path(str(getattr(uploaded_file, "name", "") or "upload.xlsx")).suffix or ".xlsx"
    target = get_saved_upload_dir() / f"last_upload{suffix}"
    data = uploaded_file.getbuffer()
    with open(target, "wb") as f:
        f.write(data)

    updated = dict(settings)
    updated["saved_upload_path"] = str(target)
    updated["saved_upload_name"] = str(getattr(uploaded_file, "name", "") or target.name)
    save_app_settings(updated)
    return updated


def clear_persisted_upload(settings: Dict[str, Any]) -> Dict[str, Any]:
    updated = dict(settings)
    existing = get_saved_upload_path(updated)
    if existing and existing.exists():
        try:
            existing.unlink()
        except Exception:
            pass
    updated["saved_upload_path"] = ""
    updated["saved_upload_name"] = ""
    save_app_settings(updated)
    return updated


def render_bottom_settings_expander(settings: Dict[str, Any]) -> Tuple[Optional[Any], Dict[str, Any]]:
    if "settings_panel_unlocked" not in st.session_state:
        st.session_state["settings_panel_unlocked"] = False

    uploaded_file = None
    current_ip = get_current_request_ip()
    ip_is_allowed = settings_visible_for_ip(settings, current_ip)
    settings_access_allowed = settings_visible(settings, current_ip)

    with st.expander("Einstellungen", expanded=False):
        st.caption("Der Bereich ist dauerhaft sichtbar und unten im UI platziert.")

        resolved_ip = str(st.session_state.get("manual_client_ip", "") or "").strip() or str(current_ip or "").strip()
        if resolved_ip and not str(st.session_state.get("manual_client_ip", "") or "").strip():
            st.session_state["manual_client_ip"] = resolved_ip

        if not settings_access_allowed:
            st.warning("Die aktuelle IP ist für die vollständigen Einstellungen nicht freigegeben. Alternativ kann die definierte Tastenkombination verwendet werden.")
            st.caption(f"Aktuell erkannte IP: {current_ip or '(nicht ermittelt)'}")
            st.text_input(
                "Test-/Override-IP",
                value=str(st.session_state.get("manual_client_ip", "") or resolved_ip),
                key="manual_client_ip",
                help="Hier kann eine andere IP eingetragen werden, um die Sichtbarkeit zu testen oder zu ändern.",
            )
            st.info("Sobald die IP-Regel passt oder die definierte Tastenkombination ausgelöst wurde, werden die übrigen Einstellungs-Tabs wieder angezeigt.")
            st.caption(f"Aktive Tastenkombination: {settings.get('shortcut_combo', '') or '(nicht gesetzt)'}")
            return uploaded_file, settings

        with st.container(border=True):
            st.markdown("### Upload")
            uploaded_file = st.file_uploader(
                str(settings.get("upload_label", "Excel-Datei hochladen")),
                type=["xlsx", "xlsm", "xls"],
                key="settings_embedded_uploader",
            )

            saved_upload_path = get_saved_upload_path(settings)
            saved_upload_name = str(settings.get("saved_upload_name", "") or "")

            if uploaded_file is not None:
                settings = persist_uploaded_file(uploaded_file, settings)
                saved_upload_path = get_saved_upload_path(settings)
                saved_upload_name = str(settings.get("saved_upload_name", "") or getattr(uploaded_file, "name", "") or "")
                st.success(f"Datei in der App gespeichert: {saved_upload_name}")

            info_cols = st.columns([3, 1])
            if saved_upload_path:
                info_cols[0].caption(f"Gespeicherte Datei: {saved_upload_name or saved_upload_path.name}")
                if info_cols[1].button("Gespeicherte Datei entfernen", key="clear_saved_upload", use_container_width=True):
                    settings = clear_persisted_upload(settings)
                    st.success("Gespeicherte Upload-Datei entfernt.")
                    st.rerun()
            else:
                info_cols[0].caption("Aktuell ist keine Upload-Datei in der App gespeichert.")

            if uploaded_file is None and saved_upload_path:
                uploaded_file = str(saved_upload_path)

        stored_hash = str(settings.get("settings_password_hash", "") or "").strip()
        protection_enabled = bool(settings.get("settings_protection_enabled", False))

        if protection_enabled and stored_hash and not st.session_state.get("settings_panel_unlocked", False):
            pw = st.text_input("Passwort für Einstellungen", type="password", key="settings_unlock_password")
            cols = st.columns([1, 1, 4])
            if cols[0].button("Freigeben", key="settings_unlock_button", use_container_width=True):
                if sha256_text(pw) == stored_hash:
                    st.session_state["settings_panel_unlocked"] = True
                    st.success("Einstellungen freigegeben.")
                    st.rerun()
                else:
                    st.error("Passwort falsch.")
            if cols[1].button("Abbrechen", key="settings_unlock_cancel", use_container_width=True):
                st.session_state["settings_panel_unlocked"] = False
                st.rerun()
            if settings.get("settings_hint"):
                st.caption(str(settings.get("settings_hint", "")))
            return uploaded_file, settings

        draft = dict(settings)
        tabs = st.tabs([
            "Header",
            "Expander",
            "Upload",
            "Schutz",
            "IP-Sichtbarkeit",
            "Tastenkombination",
            "Ansicht",
        ])

        with tabs[0]:
            draft["app_title"] = st.text_input(
                "Header",
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

        with tabs[1]:
            c1, c2 = st.columns(2)
            draft["expander_closed_bg"] = c1.color_picker(
                "Zugeklappt Hintergrund",
                value=str(settings.get("expander_closed_bg", "#e8eef8")),
                key="cfg_expander_closed_bg",
            )
            draft["expander_closed_text"] = c2.color_picker(
                "Zugeklappt Text",
                value=str(settings.get("expander_closed_text", "#0a2b55")),
                key="cfg_expander_closed_text",
            )
            c3, c4 = st.columns(2)
            draft["expander_open_bg"] = c3.color_picker(
                "Aufgeklappt Hintergrund",
                value=str(settings.get("expander_open_bg", "#0a2b55")),
                key="cfg_expander_open_bg",
            )
            draft["expander_open_text"] = c4.color_picker(
                "Aufgeklappt Text",
                value=str(settings.get("expander_open_text", "#ffffff")),
                key="cfg_expander_open_text",
            )

        with tabs[2]:
            draft["upload_label"] = st.text_input(
                "Upload-Text",
                value=str(settings.get("upload_label", "Excel-Datei hochladen")),
                key="cfg_upload_label",
            )

        with tabs[3]:
            draft["settings_protection_enabled"] = st.checkbox(
                "Schutz aktivieren",
                value=bool(settings.get("settings_protection_enabled", False)),
                key="cfg_settings_protection_enabled",
            )
            new_password = st.text_input(
                "Neues Passwort für Einstellungen",
                value="",
                type="password",
                key="cfg_settings_password",
            )
            clear_password = st.checkbox(
                "Passwortschutz entfernen",
                value=False,
                key="cfg_clear_password",
            )
            draft["settings_hint"] = st.text_input(
                "Passwort-Hinweis",
                value=str(settings.get("settings_hint", "")),
                key="cfg_settings_hint",
            )

            if not draft.get("settings_protection_enabled", False):
                draft["settings_password_hash"] = ""
            elif clear_password:
                draft["settings_password_hash"] = ""
            elif str(new_password).strip():
                draft["settings_password_hash"] = sha256_text(new_password.strip())
            else:
                draft["settings_password_hash"] = str(settings.get("settings_password_hash", "") or "")

        with tabs[4]:
            draft["settings_ip_protection_enabled"] = st.checkbox(
                "Expander nur für bestimmte IPs sichtbar",
                value=bool(settings.get("settings_ip_protection_enabled", False)),
                key="cfg_settings_ip_protection_enabled",
            )
            allowed_ips_text = st.text_area(
                "Erlaubte IP-Adressen / Bereiche",
                value="\n".join([str(x).strip() for x in settings.get("settings_allowed_ips", []) if str(x).strip()]),
                key="cfg_settings_allowed_ips",
                height=110,
                help="Eine Regel pro Zeile. Erlaubt: 192.168.1.15, 192.168.1.*, 10.0.0.0/24",
            )
            draft["settings_allowed_ips"] = parse_ip_lines(allowed_ips_text)
            invalid_ip_rules = [rule for rule in draft["settings_allowed_ips"] if not is_valid_ip_rule(rule)]
            if invalid_ip_rules:
                st.error("Ungültige IP-Regeln: " + ", ".join(invalid_ip_rules))
            st.caption(f"Aktuell erkannte IP: {current_ip or '(nicht ermittelt)'}")
            st.text_input(
                "Test-/Override-IP",
                value=str(st.session_state.get("manual_client_ip", "") or resolved_ip),
                key="manual_client_ip",
                help="Wird beim ersten Laden automatisch mit der aktuell erkannten IP vorbelegt und kann bei Bedarf manuell überschrieben werden.",
            )

        with tabs[5]:
            draft["shortcut_enabled"] = st.checkbox(
                "Tastenkombination aktiv",
                value=bool(settings.get("shortcut_enabled", False)),
                key="cfg_shortcut_enabled",
            )
            draft["shortcut_combo"] = st.text_input(
                "Tastenkombination",
                value=str(settings.get("shortcut_combo", "")),
                key="cfg_shortcut_combo",
                help="Beispiel: ctrl+alt+s. Kann zusätzlich zur IP-Sichtbarkeit als Freischaltung genutzt werden.",
            )

        with tabs[6]:
            st.caption(f"Einstellungsdatei: {SETTINGS_FILE}")
            st.write("Dieser Expander bleibt unten im UI sichtbar und kann direkt geöffnet werden, sofern die IP-Regel passt.")

        c1, c2, c3 = st.columns([1, 2, 2])
        if c1.button("Einstellungen speichern", key="settings_save_button", use_container_width=True):
            invalid_ip_rules = [rule for rule in draft.get("settings_allowed_ips", []) if not is_valid_ip_rule(rule)]
            if invalid_ip_rules:
                st.error("Bitte zuerst die ungültigen IP-Regeln korrigieren.")
            else:
                save_app_settings(draft)
                st.success("Einstellungen gespeichert.")
                st.rerun()
        if c2.button("Schutzstatus zurücksetzen", key="settings_reset_unlock", use_container_width=True):
            st.session_state["settings_panel_unlocked"] = False
            st.info("Freigabe zurückgesetzt.")
            st.rerun()
        if c3.button("Tastenkombination-Freigabe zurücksetzen", key="settings_reset_shortcut_override", use_container_width=True):
            st.session_state["settings_shortcut_override"] = False
            st.info("Tastenkombination-Freigabe zurückgesetzt.")
            st.rerun()

    return uploaded_file, settings


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
        return None

    value = normalize_text(answer_value)
    for option in question.options:
        if option.text == value:
            return option.next_question
    return None


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
        for option in questions[q_no].options:
            if option.next_question:
                stack.append(option.next_question)
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
    input_type = normalize_key(question.input_type)
    is_number = "zahl" in input_type or "anzahl" in normalize_key(question.text) or "wie alt" in normalize_key(question.text)

    if is_number:
        value = st.number_input("Antwort eingeben", step=1, key=f"input_{question.number}")
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
register_settings_shortcut(app_settings)
sync_shortcut_unlock()

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f6f8fc;
    }
    .block-container {
        max-width: 100% !important;
        padding-top: 0.4rem;
        padding-left: 1.6rem;
        padding-right: 1.6rem;
        padding-bottom: 2rem;
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
    details > summary {
        background: {app_settings.get("expander_closed_bg", "#e8eef8")} !important;
        color: {app_settings.get("expander_closed_text", "#0a2b55")} !important;
        border-radius: 0.8rem !important;
        border: 1px solid rgba(10, 43, 85, 0.14) !important;
    }
    details[open] > summary {
        background: {app_settings.get("expander_open_bg", "#0a2b55")} !important;
        color: {app_settings.get("expander_open_text", "#ffffff")} !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if app_settings.get("show_header", True):
    st.title(str(app_settings.get("app_title", "Dynamischer Frage-Antwort-Mechanismus")))
if app_settings.get("show_subheader", True):
    st.write(str(app_settings.get("app_subtitle", "Excel-Datei hochladen, Fragenpfad automatisch durchlaufen und Antworten am Ende als CSV oder Excel exportieren.")))

uploaded_file = None
workbook_label = ""

st.markdown('<div style="height: 1rem;"></div>', unsafe_allow_html=True)
uploaded_file, app_settings = render_bottom_settings_expander(app_settings)

if uploaded_file is not None:
    if isinstance(uploaded_file, str):
        workbook_label = Path(uploaded_file).name
    else:
        workbook_label = getattr(uploaded_file, "name", "") or ""

if uploaded_file is not None:
    workbook_label = workbook_label if "workbook_label" in locals() else ""
    reload_needed = st.session_state.uploaded_filename != workbook_label or st.session_state.questions is None
    if reload_needed:
        try:
            questions, sheet_name = load_questions_from_excel(uploaded_file)
            st.session_state.questions = questions
            st.session_state.uploaded_filename = workbook_label
            st.session_state.sheet_name = sheet_name
            reset_questionnaire(keep_questions=True)
            st.success(
                f"Datei '{workbook_label}' geladen. {len(questions)} Fragen erkannt. Verwendetes Blatt: {sheet_name}"
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
