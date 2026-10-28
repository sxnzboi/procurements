"""Backward-compatible storage helpers for daily meals in metadata.notes."""

import json

_PREFIX = "__skdj_daily_meals_v1__:"
_DEFAULTS = {
    "breakfast_menu": "",
    "lunch_menu": "",
    "dinner_menu": "",
    "child_count": "",
}


def decode_daily_notes(raw_notes):
    raw_notes = raw_notes or ""
    if raw_notes.startswith(_PREFIX):
        try:
            payload = json.loads(raw_notes[len(_PREFIX):])
            result = {"notes": payload.get("notes", "") or ""}
            result.update({key: str(payload.get(key, "") or "") for key in _DEFAULTS})
            return result
        except (TypeError, ValueError):
            pass
    return {"notes": raw_notes, **_DEFAULTS}


def encode_daily_notes(raw_notes, *, notes=None, meal_plan=None):
    payload = decode_daily_notes(raw_notes)
    if notes is not None:
        payload["notes"] = str(notes).strip()
    if meal_plan:
        for key in _DEFAULTS:
            if key in meal_plan:
                payload[key] = str(meal_plan[key] if meal_plan[key] is not None else "").strip()

    if not any(payload[key] for key in _DEFAULTS):
        return payload["notes"]
    encoded = {"notes": payload["notes"], **{key: payload[key] for key in _DEFAULTS}}
    return _PREFIX + json.dumps(encoded, ensure_ascii=False, separators=(",", ":"))


def decode_daily_metadata(metadata):
    result = dict(metadata)
    decoded = decode_daily_notes(result.get("notes"))
    result["notes"] = decoded["notes"]
    for key in _DEFAULTS:
        result[key] = decoded[key]
    return result
