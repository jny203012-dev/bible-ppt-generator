import json
import re

# 기존 JSON 불러오기
try:
    with open('bible.json', 'r', encoding='utf-8') as f:
        result = json.load(f)
except:
    result = {}

with open('bible.txt', 'r', encoding='utf-8') as f:
    full_text = f.read()

# 책 이름 추출 (첫 줄)
lines = full_text.split('\n')
current_book = lines[0].strip()

# 전체 텍스트에서 "숫자:숫자" 기준으로 나누기
matches = re.split(r'(\d+:\d+)', full_text)

# matches 구조:
# ['', '1:1', ' 내용...', '1:2', ' 내용...', ...]

for i in range(1, len(matches), 2):
    verse_part = matches[i]
    text = matches[i + 1].strip().replace('\n', ' ')

    key = f"{current_book} {verse_part}"
    result[key] = text

# 저장
with open('bible.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("숫자 기준 분리 완료!")