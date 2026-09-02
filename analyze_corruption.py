import re

with open('index.html', 'r', encoding='utf-8') as f:
    text = f.read()

FFFD = '\ufffd'
positions = [m.start() for m in re.finditer(FFFD, text)]
print('Total FFFD:', len(positions))

with open('fffd_contexts.txt', 'w', encoding='utf-8') as out:
    for i, pos in enumerate(positions):
        start = max(0, pos - 20)
        end = min(len(text), pos + 20)
        snippet = text[start:end].replace('\n', '\\n')
        out.write(f'{i}\t{pos}\t{snippet}\n')

# Also find suspicious lone "?" characters (surrounded by non-alphanumeric, likely corrupted arrows)
# Pattern: space-?-space near a percentage number
q_positions = []
for m in re.finditer(r'(?<=[ >])\?(?=[ %])', text):
    q_positions.append(m.start())
print('Suspicious ? count:', len(q_positions))
with open('q_contexts.txt', 'w', encoding='utf-8') as out:
    for i, pos in enumerate(q_positions):
        start = max(0, pos - 25)
        end = min(len(text), pos + 25)
        snippet = text[start:end].replace('\n', '\\n')
        out.write(f'{i}\t{pos}\t{snippet}\n')
