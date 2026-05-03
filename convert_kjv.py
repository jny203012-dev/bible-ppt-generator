import json
import re

result = {}

with open('bible_kjv.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for line in lines:
    line = line.strip()
    if not line:
        continue

    # 1. "딤전6:2 ..." 분리
    match = re.match(r'(\D+)(\d+):(\d+)\s+(.*)', line)
    
    if not match:
        continue

    book = match.group(1)
    chapter = match.group(2)
    verse = match.group(3)
    text = match.group(4)

    # 2. <제목> 제거
    text = re.sub(r'<.*?>', '', text).strip()

    # 3. key 생성 (공백 추가)
    key = f"{book} {chapter}:{verse}"

    result[key] = text

# 4. JSON 저장
with open('bible_kor_revised.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("변환 완료!")