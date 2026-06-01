import json
import re

BOOKS = [
    "창", "출", "레", "민", "신", "수", "삿", "룻",
    "삼상", "삼하", "왕상", "왕하", "대상", "대하",
    "스", "느", "에", "욥", "시", "잠", "전", "아",
    "사", "렘", "애", "겔", "단", "호", "욜", "암",
    "옵", "욘", "미", "나", "합", "습", "학", "슥", "말",
    "마", "막", "눅", "요", "행", "롬", "고전", "고후",
    "갈", "엡", "빌", "골", "살전", "살후", "딤전", "딤후",
    "딛", "몬", "히", "약", "벧전", "벧후", "요일", "요이",
    "요삼", "유", "계"
]

book_pattern = "|".join(sorted(BOOKS, key=len, reverse=True))
embedded_ref_pattern = re.compile(rf"({book_pattern})\s*(\d+):(\d+)\s*")

with open("bible.json", "r", encoding="utf-8") as f:
    bible_data = json.load(f)

fixed_data = bible_data.copy()

target_version = "개역개정"

verses = fixed_data[target_version]

for key, text in list(verses.items()):
    match = embedded_ref_pattern.search(text)

    if match:
        before = text[:match.start()].strip()
        book = match.group(1)
        chapter = match.group(2)
        verse = match.group(3)
        after = text[match.end():].strip()

        next_key = f"{book} {chapter}:{verse}"

        print(f"수정: {key} 안에서 {next_key} 발견")

        verses[key] = before

        if next_key not in verses or not verses[next_key].strip():
            verses[next_key] = after

with open("bible_fixed.json", "w", encoding="utf-8") as f:
    json.dump(fixed_data, f, ensure_ascii=False, indent=2)

print("완료: bible_fixed.json 생성됨")