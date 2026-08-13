from collections import OrderedDict
from dataclasses import dataclass, field
import re


@dataclass
class HardwareItem:
    product: str
    model: str = ""
    serials: list[str] = field(default_factory=list)


def parse_pasted_values(text: str) -> list[str]:
    if not text or not text.strip():
        return []
    pieces = re.split(r"[\n\r\t,;]+", text)
    return [piece.strip() for piece in pieces if piece.strip()]


def find_duplicates(values) -> list[str]:
    seen = {}
    duplicates = []
    duplicate_keys = set()
    for raw in values:
        value = raw.strip()
        if not value:
            continue
        key = value.casefold()
        if key in seen and key not in duplicate_keys:
            duplicates.append(seen[key])
            duplicate_keys.add(key)
        else:
            seen.setdefault(key, value)
    return duplicates


def parse_bulk_hardware(text: str):
    grouped = OrderedDict()
    ignored = []
    for raw_line in text.splitlines():
        if not raw_line.strip():
            continue
        parts = [part.strip() for part in raw_line.split("\t") if part.strip()]
        if len(parts) >= 3:
            product, model = parts[0], parts[1]
            serial_text = "\t".join(parts[2:])
        elif len(parts) == 2:
            product, model, serial_text = parts[0], "", parts[1]
        else:
            ignored.append(raw_line)
            continue
        serials = parse_pasted_values(serial_text)
        if not product or not serials:
            ignored.append(raw_line)
            continue
        grouped.setdefault((product, model), []).extend(serials)
    items = [HardwareItem(product=p, model=m, serials=s) for (p, m), s in grouped.items()]
    return items, ignored


def hardware_serial_count(items) -> int:
    return sum(len(item.serials) for item in items)


def safe_filename_part(value: str, fallback: str = "Report") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("_.") or fallback
