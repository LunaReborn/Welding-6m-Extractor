import re


def clean_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def chunk_text(text: str, max_chars: int = 1500, overlap: int = 120) -> list[str]:
    text = clean_text(text)
    if not text:
        return []
    if len(text) <= max_chars:
        return [text]

    paragraphs = [p.strip() for p in re.split(r"\n{2,}|(?<=[。；;])\s*", text) if p.strip()]
    chunks: list[str] = []
    current = ""

    for para in paragraphs:
        if len(para) > max_chars:
            if current:
                chunks.append(current.strip())
                current = ""
            for i in range(0, len(para), max_chars - overlap):
                chunks.append(para[i : i + max_chars].strip())
            continue

        candidate = f"{current}\n{para}" if current else para
        if len(candidate) <= max_chars:
            current = candidate
        else:
            chunks.append(current.strip())
            prefix = current[-overlap:] if overlap > 0 else ""
            current = f"{prefix}\n{para}" if prefix else para

    if current.strip():
        chunks.append(current.strip())
    return chunks
