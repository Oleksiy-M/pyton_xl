from translate import Translator

# Ініціалізуємо перекладач з англійської на російську
translator = Translator(from_lang="en", to_lang="ru")

# Відкриваємо файл з субтитрами
with open("Kid_eng.srt", "r", encoding="utf-8") as file:
    lines = file.readlines()

translated_lines = []

# Обробляємо кожен рядок
for line in lines:
    # Якщо рядок містить тільки цифри або час — не перекладаємо
    if line.strip().isdigit() or "-->" in line:
        translated_lines.append(line)
    elif line.strip() == "":
        translated_lines.append(line)
    else:
        # Перекладаємо текст
        translated_text = translator.translate(line.strip())
        translated_lines.append(translated_text + "\n")

# Записуємо результат у новий файл
with open("subtitles_ru.srt", "w", encoding="utf-8") as file:
    file.writelines(translated_lines)

print("Переклад завершено. Збережено в subtitles_ru.srt")

