import re

def parse_srt(filename):
    """Зчитати субтитри у форматі SRT у список словників"""
    subs = []
    with open(filename, encoding='utf-8') as f:
        content = f.read().strip()

    # Розбиваємо на блоки субтитрів
    blocks = re.split(r'\n\n+', content)
    for block in blocks:
        lines = block.split('\n')
        if len(lines) >= 3:
            index = lines[0].strip()
            time_line = lines[1].strip()
            text_lines = lines[2:]
            start, end = time_line.split(' --> ')
            subs.append({
                'index': int(index),
                'start': start,
                'end': end,
                'text': '\n'.join(text_lines)
            })
    return subs

def time_to_ms(t):
    """Перетворити час 'HH:MM:SS,mmm' у мілісекунди для зручного порівняння"""
    h, m, s = t.split(':')
    s, ms = s.split(',')
    return (int(h)*3600 + int(m)*60 + int(s))*1000 + int(ms)

def merge_subs(main_subs, add_subs):
    """Об’єднати два списки субтитрів, уникаючи накладань"""
    merged = main_subs[:]

    # Ітеруємося по доповнюючих субтитрах
    for add in add_subs:
        add_start = time_to_ms(add['start'])
        add_end = time_to_ms(add['end'])
        # Перевірка, чи немає перекриття з існуючими субтитрами
        overlap = False
        for main in main_subs:
            main_start = time_to_ms(main['start'])
            main_end = time_to_ms(main['end'])
            # Якщо є хоча б невелике накладання часу - пропускаємо додавання
            if not (add_end < main_start or add_start > main_end):
                overlap = True
                break
        if not overlap:
            merged.append(add)

    # Відсортувати за часом початку
    merged.sort(key=lambda x: time_to_ms(x['start']))

    # Перенумерувати індекси
    for i, sub in enumerate(merged, 1):
        sub['index'] = i

    return merged

def write_srt(subs, filename):
    """Записати список субтитрів у SRT файл"""
    with open(filename, 'w', encoding='utf-8') as f:
        for sub in subs:
            f.write(str(sub['index']) + '\n')
            f.write(f"{sub['start']} --> {sub['end']}\n")
            f.write(sub['text'] + '\n\n')

if __name__ == '__main__':
    main_file = 'KarateKid.srt'
    add_file = 'KarateKid_ru.srt'
    output_file = 'merged.srt'

    main_subs = parse_srt(main_file)
    add_subs = parse_srt(add_file)

    merged_subs = merge_subs(main_subs, add_subs)
    write_srt(merged_subs, output_file)

    print(f"Об’єднання завершено. Файл '{output_file}' створено.")
