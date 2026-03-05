from __future__ import annotations

import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Dict, List, Tuple

import cv2
import pytesseract


CANONICAL_FIELDS: List[str] = [
    "Power",
    "Merits",
    "Units Killed",
    "Units Dead",
    "Units Healed",
    "Total Resources Gathered",
    "Gold Gathered",
    "Wood Gathered",
    "Ore Gathered",
    "Mana Gathered",
    "Gems Gathered",
    "Total Resource Assistance Given",
    "Times Resource Assistance Given",
    "Times Alliance Help Given",
]


VALUE_RE = re.compile(r"^\d[\d,]*$")


@dataclass
class OCRLine:
    text: str
    words: List[str]
    left: int
    top: int
    right: int
    bottom: int


def _configure_tesseract() -> None:
    env_cmd = os.environ.get("TESSERACT_CMD")
    if env_cmd:
        pytesseract.pytesseract.tesseract_cmd = env_cmd
        return

    default_windows_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default_windows_cmd):
        pytesseract.pytesseract.tesseract_cmd = default_windows_cmd


def _normalize_label(text: str) -> str:
    text = re.sub(r"\b0([a-zA-Z]+)\b", r"o\1", text)
    cleaned = re.sub(r"[^a-zA-Z ]+", " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().lower()
    return cleaned


def _parse_numeric(token: str) -> int | None:
    token = token.strip()
    token = re.sub(r"[^\d,]", "", token)
    if not token or not VALUE_RE.match(token):
        return None
    return int(token.replace(",", ""))


def _has_alpha(token: str) -> bool:
    return any(ch.isalpha() for ch in token)


def _token_as_label(token: str) -> str:
    token = re.sub(r"\b0([a-zA-Z]+)\b", r"O\1", token)
    token = re.sub(r"[^0-9a-zA-Z]", "", token)
    return token


def _line_to_label_value(words: List[str]) -> Tuple[str, int] | None:
    if not words:
        return None

    for idx in range(len(words) - 1, -1, -1):
        value = _parse_numeric(words[idx])
        if value is None:
            continue

        raw_label_tokens = words[:idx]
        label_tokens = []
        for tok in raw_label_tokens:
            if _parse_numeric(tok) is not None and not _has_alpha(tok):
                continue
            cleaned = _token_as_label(tok)
            if cleaned and _has_alpha(cleaned):
                label_tokens.append(cleaned)

        if not label_tokens:
            continue
        label = " ".join(label_tokens).strip()
        if label:
            return (label, value)
    return None


def _line_label_only(words: List[str]) -> str | None:
    tokens: List[str] = []
    for tok in words:
        if _parse_numeric(tok) is not None and not _has_alpha(tok):
            continue
        cleaned = _token_as_label(tok)
        if cleaned and _has_alpha(cleaned):
            tokens.append(cleaned)
    if not tokens:
        return None
    return " ".join(tokens)


def _line_numeric_only(words: List[str]) -> int | None:
    if not words:
        return None
    if any(_has_alpha(tok) for tok in words):
        return None
    digit_chunks: List[str] = []
    for tok in words:
        cleaned = re.sub(r"[^\d]", "", tok)
        if cleaned:
            digit_chunks.append(cleaned)
    if digit_chunks:
        return int("".join(digit_chunks))
    return None


def _extract_number_candidates(text: str) -> List[int]:
    candidates: List[int] = []
    for match in re.finditer(r"\d[\d,\s]{1,}", text):
        digits = re.sub(r"[^\d]", "", match.group(0))
        if len(digits) >= 3:
            candidates.append(int(digits))
    return candidates


def _best_single_label_match(label: str, target: str) -> float:
    return SequenceMatcher(None, _normalize_label(label), _normalize_label(target)).ratio()


def _extract_power_merits_fallback(image_path: str) -> Dict[str, int]:
    image = cv2.imread(image_path)
    if image is None:
        return {}

    h, w = image.shape[:2]
    left_crop = image[:, : max(1, int(w * 0.38))]
    left_crop = cv2.resize(left_crop, None, fx=1.8, fy=1.8, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(left_crop, cv2.COLOR_BGR2GRAY)
    th = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 15
    )

    variants = [gray, th]
    configs = ["--oem 3 --psm 6", "--oem 3 --psm 4", "--oem 3 --psm 11"]
    targets = ["Power", "Merits"]
    found: Dict[str, int] = {}

    for variant in variants:
        for cfg in configs:
            text = pytesseract.image_to_string(variant, config=cfg)
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines:
                continue

            for idx, line in enumerate(lines):
                for target in targets:
                    if target in found:
                        continue
                    score = _best_single_label_match(line, target)
                    if score < 0.58:
                        continue

                    # Same-line number first.
                    same_line_numbers = _extract_number_candidates(line)
                    if same_line_numbers:
                        found[target] = max(same_line_numbers, key=lambda x: len(str(x)))
                        continue

                    # Then inspect next few lines for stacked value.
                    for j in range(idx + 1, min(len(lines), idx + 5)):
                        next_numbers = _extract_number_candidates(lines[j])
                        if next_numbers:
                            found[target] = max(next_numbers, key=lambda x: len(str(x)))
                            break

            if "Power" in found and "Merits" in found:
                return found

    return found


def _best_field_match(label: str, candidates: List[str]) -> Tuple[str, float]:
    normalized = _normalize_label(label)
    best_name = ""
    best_score = 0.0
    for candidate in candidates:
        score = SequenceMatcher(None, normalized, _normalize_label(candidate)).ratio()
        if score > best_score:
            best_name = candidate
            best_score = score
    return best_name, best_score


def _preprocess_image(image_path: str):
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError(f"Unable to read image: {image_path}")

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(gray, h=25)
    thresh = cv2.adaptiveThreshold(
        denoised,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        35,
        15,
    )
    return thresh


def _extract_lines(image) -> List[OCRLine]:
    data = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DICT,
        config="--oem 3 --psm 6",
    )

    buckets: Dict[Tuple[int, int, int, int], List[Tuple[int, int, int, int, str]]] = {}
    total = len(data["text"])
    for idx in range(total):
        text = data["text"][idx].strip()
        if not text:
            continue

        try:
            conf = float(data["conf"][idx])
        except ValueError:
            continue
        if conf < 15:
            continue

        key = (
            data["page_num"][idx],
            data["block_num"][idx],
            data["par_num"][idx],
            data["line_num"][idx],
        )
        left = int(data["left"][idx])
        top = int(data["top"][idx])
        width = int(data["width"][idx])
        height = int(data["height"][idx])
        buckets.setdefault(key, []).append((left, top, left + width, top + height, text))

    lines: List[OCRLine] = []
    for entries in buckets.values():
        entries.sort(key=lambda x: x[0])
        words = [item[4] for item in entries]
        left = min(item[0] for item in entries)
        top = min(item[1] for item in entries)
        right = max(item[2] for item in entries)
        bottom = max(item[3] for item in entries)
        lines.append(
            OCRLine(
                text=" ".join(words),
                words=words,
                left=left,
                top=top,
                right=right,
                bottom=bottom,
            )
        )

    lines.sort(key=lambda ln: (ln.top, ln.left))
    return lines


def _normalize_anchor(text: str) -> str:
    return re.sub(r"[^a-z]", "", text.lower())


def _is_lord_like(text: str) -> bool:
    normalized = _normalize_anchor(text)
    if not normalized:
        return False
    if "lord" in normalized:
        return True
    return SequenceMatcher(None, normalized, "lord").ratio() >= 0.72


def _clean_nickname(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", text).strip()
    cleaned = cleaned.strip("|:;,-_ ")
    return cleaned


def _sanitize_nickname_candidate(text: str) -> str:
    candidate = _clean_nickname(text)
    if not candidate:
        return ""

    # Remove leading "Lord" marker if present.
    candidate = re.sub(r"^lord\b[:\-\s]*", "", candidate, flags=re.IGNORECASE).strip()

    # If OCR merged left profile text with right stat column, cut at stat field phrase.
    lowered = candidate.lower()
    cut_idx = len(candidate)
    for field in CANONICAL_FIELDS:
        idx = lowered.find(field.lower())
        if idx != -1 and idx < cut_idx:
            cut_idx = idx
    candidate = candidate[:cut_idx]

    # Remove trailing numeric chunk from OCR merges like "<name> 123,456".
    candidate = re.sub(r"\s*\d[\d,\s]*$", "", candidate)
    return _clean_nickname(candidate)


def _is_likely_stat_label(text: str) -> bool:
    _, score = _best_field_match(text, CANONICAL_FIELDS)
    return score >= 0.74


def _is_unlikely_nickname(text: str) -> bool:
    normalized = _normalize_label(text)
    if not normalized:
        return True

    words = normalized.split()
    if len(words) > 4:
        return True

    blocked = {
        "power",
        "merits",
        "resources",
        "resource",
        "info",
        "units",
        "killed",
        "dead",
        "healed",
        "gathered",
        "assistance",
        "times",
        "alliance",
        "help",
        "more",
        "total",
        "gold",
        "wood",
        "ore",
        "mana",
        "gems",
    }
    if any(word in blocked for word in words):
        return True
    return False


def _extract_nickname_from_lines(lines: List[OCRLine]) -> str | None:
    for idx, line in enumerate(lines):
        line_text = line.text.strip()
        if not line_text:
            continue
        if not _is_lord_like(line_text):
            continue

        # Same-line fallback if OCR merged "Lord <nickname>" into one line.
        for w_idx, word in enumerate(line.words):
            if not _is_lord_like(word):
                continue
            same_line_nick = _sanitize_nickname_candidate(" ".join(line.words[w_idx + 1 :]))
            if same_line_nick and same_line_nick.lower() != "lord":
                if _is_likely_stat_label(same_line_nick):
                    continue
                if _is_unlikely_nickname(same_line_nick):
                    continue
                return same_line_nick

        # Primary path: nickname is rendered on the line(s) directly below "Lord".
        for next_line in lines[idx + 1 : idx + 8]:
            dy = next_line.top - line.bottom
            # OCR boxes can slightly overlap even when text is visually below.
            if dy < -35:
                continue
            if dy > 220:
                break
            if next_line.right < line.left - 120:
                continue
            if next_line.left > line.right + 260:
                continue

            candidate = _sanitize_nickname_candidate(next_line.text)
            if not candidate:
                continue
            if _is_lord_like(candidate):
                continue
            if _is_likely_stat_label(candidate):
                continue
            if _is_unlikely_nickname(candidate):
                continue
            if _line_numeric_only(next_line.words) is not None:
                continue
            if not any(ch.isalpha() for ch in candidate):
                continue
            return candidate

    return None


def _find_power_anchor_line(lines: List[OCRLine]) -> OCRLine | None:
    best_line = None
    best_score = 0.0
    for line in lines:
        score = _best_single_label_match(line.text, "Power")
        for word in line.words:
            word_score = _best_single_label_match(word, "Power")
            if word_score > score:
                score = word_score
        if score > best_score:
            best_score = score
            best_line = line
    if best_score >= 0.72:
        return best_line
    return None


def _extract_nickname_from_power_anchor(lines: List[OCRLine]) -> str | None:
    power_line = _find_power_anchor_line(lines)
    if power_line is None:
        return None

    candidates: List[Tuple[float, str]] = []
    for line in lines:
        if line.bottom >= power_line.top:
            continue
        if power_line.top - line.bottom > 280:
            continue
        if line.right < power_line.left - 140:
            continue
        if line.left > power_line.right + 260:
            continue

        candidate = _sanitize_nickname_candidate(line.text)
        if not candidate:
            continue
        if _is_lord_like(candidate):
            continue
        if _is_likely_stat_label(candidate):
            continue
        if _is_unlikely_nickname(candidate):
            continue
        if _line_numeric_only(line.words) is not None:
            continue
        if not any(ch.isalpha() for ch in candidate):
            continue

        distance_score = 1.0 - ((power_line.top - line.bottom) / 281.0)
        text_len_score = min(1.0, len(candidate) / 16.0)
        score = (distance_score * 0.7) + (text_len_score * 0.3)
        candidates.append((score, candidate))

    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1]


def _extract_nickname_from_roi(image_path: str, lines: List[OCRLine]) -> str | None:
    power_line = _find_power_anchor_line(lines)
    if power_line is None:
        return None

    image = cv2.imread(image_path)
    if image is None:
        return None

    h, w = image.shape[:2]
    x1 = max(0, power_line.left - 150)
    x2 = min(w, power_line.left + 560)
    y1 = max(0, power_line.top - 360)
    y2 = max(0, min(h, power_line.top - 10))
    if x2 - x1 < 80 or y2 - y1 < 50:
        return None

    roi = image[y1:y2, x1:x2]
    roi = cv2.resize(roi, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    th = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 12
    )

    variants = [gray, th]
    configs = ["--oem 3 --psm 6", "--oem 3 --psm 11"]

    for variant in variants:
        for cfg in configs:
            text = pytesseract.image_to_string(variant, config=cfg)
            lines_text = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines_text:
                continue

            for idx, raw in enumerate(lines_text):
                if not _is_lord_like(raw):
                    continue

                # If OCR merges "Lord Saint Angel" on one line, keep text after Lord.
                words = raw.split()
                for w_idx, word in enumerate(words):
                    if not _is_lord_like(word):
                        continue
                    merged = _sanitize_nickname_candidate(" ".join(words[w_idx + 1 :]))
                    if merged and merged.lower() != "lord":
                        if _is_likely_stat_label(merged):
                            continue
                        if _is_unlikely_nickname(merged):
                            continue
                        return merged

                # Otherwise take the first clean line below "Lord".
                for next_raw in lines_text[idx + 1 : idx + 5]:
                    candidate = _sanitize_nickname_candidate(next_raw)
                    if not candidate:
                        continue
                    if _is_lord_like(candidate):
                        continue
                    if _is_likely_stat_label(candidate):
                        continue
                    if _is_unlikely_nickname(candidate):
                        continue
                    if _parse_numeric(candidate) is not None:
                        continue
                    if not any(ch.isalpha() for ch in candidate):
                        continue
                    return candidate

            # Fallback: choose strongest name-like line in ROI when "Lord" was missed.
            roi_candidates: List[str] = []
            for raw in lines_text:
                candidate = _sanitize_nickname_candidate(raw)
                if not candidate:
                    continue
                if _is_lord_like(candidate):
                    continue
                if _is_likely_stat_label(candidate):
                    continue
                if _is_unlikely_nickname(candidate):
                    continue
                if _parse_numeric(candidate) is not None:
                    continue
                if not any(ch.isalpha() for ch in candidate):
                    continue
                roi_candidates.append(candidate)
            if roi_candidates:
                roi_candidates.sort(key=lambda s: len(s), reverse=True)
                return roi_candidates[0]

    return None


def _extract_stats_from_lines(lines: List[OCRLine], image_path: str) -> Dict[str, int]:

    # Keep the best value per field by quality score.
    result: Dict[str, int] = {}
    quality_by_field: Dict[str, float] = {}

    def assign(field: str, value: int, quality: float) -> None:
        if field not in quality_by_field or quality > quality_by_field[field]:
            result[field] = value
            quality_by_field[field] = quality

    # Pass 1: inline "Label ... Value" lines.
    for line in lines:
        pair = _line_to_label_value(line.words)
        if not pair:
            continue

        label, value = pair
        match, score = _best_field_match(label, CANONICAL_FIELDS)
        if score < 0.62:
            continue

        assign(match, value, score + 1.0)

    # Pass 2: label-only + numeric-only line pairing.
    label_lines: Dict[str, OCRLine] = {}
    numeric_lines: List[Tuple[int, OCRLine, int]] = []

    for idx, line in enumerate(lines):
        value = _line_numeric_only(line.words)
        if value is not None:
            numeric_lines.append((idx, line, value))
            continue

        label = _line_label_only(line.words)
        if not label:
            continue
        match, score = _best_field_match(label, CANONICAL_FIELDS)
        if score < 0.60:
            continue
        if match not in label_lines or line.top < label_lines[match].top:
            label_lines[match] = line

    used_numeric_idxs = set()
    # Process top-to-bottom so nearby values map in visual order.
    ordered_missing = sorted(
        [field for field in CANONICAL_FIELDS if field not in result and field in label_lines],
        key=lambda f: label_lines[f].top,
    )

    for field in ordered_missing:
        label_line = label_lines[field]
        label_cy = (label_line.top + label_line.bottom) / 2.0
        label_h = max(1.0, float(label_line.bottom - label_line.top))
        best_candidate = None
        best_candidate_score = -1.0

        for idx, num_line, num_value in numeric_lines:
            if idx in used_numeric_idxs:
                continue

            num_cy = (num_line.top + num_line.bottom) / 2.0
            dy = abs(label_cy - num_cy)

            # Case A: same row, value on the right column.
            if num_line.left >= label_line.right - 20:
                same_row_limit = max(18.0, label_h * 0.9)
                if dy <= same_row_limit:
                    score = 2.0 - (dy / (same_row_limit + 1.0))
                    if score > best_candidate_score:
                        best_candidate = (idx, num_value, score + 0.8)
                        best_candidate_score = score

            # Case B: value on next line below (Power/Merits style).
            dy_down = num_line.top - label_line.bottom
            if 0 <= dy_down <= 120:
                x_close = abs(num_line.left - label_line.left) <= 180
                x_overlap = not (num_line.right < label_line.left or num_line.left > label_line.right + 220)
                if x_close or x_overlap:
                    score = 1.6 - (dy_down / 121.0)
                    if score > best_candidate_score:
                        best_candidate = (idx, num_value, score + 0.4)
                        best_candidate_score = score

        if best_candidate is not None:
            idx, value, qual = best_candidate
            assign(field, value, qual)
            used_numeric_idxs.add(idx)

    # Dedicated fallback for left-panel stacked fields.
    if "Power" not in result or "Merits" not in result:
        fallback = _extract_power_merits_fallback(image_path)
        for key in ("Power", "Merits"):
            if key in fallback:
                assign(key, fallback[key], 10.0)

    return result


def extract_stats_and_nickname(image_path: str) -> Tuple[Dict[str, int], str | None]:
    _configure_tesseract()
    preprocessed = _preprocess_image(image_path)
    lines = _extract_lines(preprocessed)
    stats = _extract_stats_from_lines(lines, image_path)
    nickname = _extract_nickname_from_lines(lines)
    if not nickname:
        nickname = _extract_nickname_from_power_anchor(lines)
    if not nickname:
        nickname = _extract_nickname_from_roi(image_path, lines)
    return stats, nickname


def extract_stats(image_path: str) -> Dict[str, int]:
    stats, _ = extract_stats_and_nickname(image_path)
    return stats


def extract_nickname(image_path: str) -> str | None:
    _, nickname = extract_stats_and_nickname(image_path)
    return nickname
