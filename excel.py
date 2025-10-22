import os
import tempfile
import logging
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# === Настройки из переменных окружения ===
BOT_TOKEN = os.environ["BOT_TOKEN"]
AUTHORIZED_IDS_STR = os.environ.get("AUTHORIZED_IDS", "")
AUTHORIZED_USER_IDS = set(map(int, AUTHORIZED_IDS_STR.split(","))) if AUTHORIZED_IDS_STR else set()

# === Список приоритетных напитков ===
PRIORITY_DRINKS = {
    "Espresso",
    "Double espresso decaffeinated",
    "Chocolate Truffle",
    "Sakura Latte",
    "Matcha Latte",
    "Berry RAF",
    "Kakao Banana",
    "Masala Tea Latte",
    "Cheese & Orange Latte",
    "Double cappuccino vegan",
    "Flat White",
    "Flat White decaffeinated",
    "Flat white vegan",
    "Latte",
    "Latte decaffeinated",
    "Latte vegan",
    "Ice latte",
    "Ice latte decaffeinated",
    "Espresso decaffeinated",
    "Ice latte vegan",
    "Espresso tonic",
    "Espresso tonic decaffeinated",
    "Bumblebee",
    "Tea",
    "Doppio(double espresso)",
    "Americano",
    "Americano decaffeinated",
    "Cappuccino",
    "Cappuccino decaffeinated",
    "Cacao",
    "Hot chocolate",
    "Cappuccino vegan",
    "Double Americano",
    "Double cappuccino"
}
PRIORITY_DRINKS_LOWER = {name.lower().strip() for name in PRIORITY_DRINKS}


def is_authorized(user_id: int) -> bool:
    return not AUTHORIZED_USER_IDS or user_id in AUTHORIZED_USER_IDS


def format_number(x):
    """Преобразует 43.0 → '43', 43.5 → '43.5'"""
    s = f"{x:.2f}".rstrip('0').rstrip('.')
    return s if s != '' else '0'


def analyze_excel(file_path: str) -> tuple[str, str, pd.DataFrame]:
    df_raw = pd.read_excel(file_path, header=None)

    # Найти строку с заголовками
    header_row = None
    for i in range(len(df_raw)):
        if "Denumire marfa" in df_raw.iloc[i].values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("❌ Не найдены заголовки. Убедитесь, что файл — отчёт кассы.")

    df = df_raw.iloc[header_row:].copy()
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # Извлечь дату из столбца 'Data'
    report_date = "неизвестна"
    if 'Data' in df.columns:
        non_empty = df['Data'].dropna()
        if not non_empty.empty:
            try:
                report_date = pd.to_datetime(non_empty.iloc[0], dayfirst=True).strftime('%d.%m.%Y')
            except Exception:
                report_date = str(non_empty.iloc[0]).strip()

    # Проверка наличия нужных столбцов
    required_cols = ["Denumire marfa", "Cantitate", "Suma cu TVA fără reducere"]
    if not all(col in df.columns for col in required_cols):
        raise ValueError("❌ Отсутствуют необходимые столбцы.")

    df = df[required_cols].copy()
    df = df.dropna(subset=["Denumire marfa"])
    df = df[~df["Denumire marfa"].str.contains("Punga", na=False)]

    # Пометить приоритетные напитки
    df['is_priority'] = df['Denumire marfa'].str.lower().str.strip().isin(PRIORITY_DRINKS_LOWER)

    # Агрегация
    result = df.groupby("Denumire marfa").agg(
        Количество=("Cantitate", "sum"),
        Сумма=("Suma cu TVA fără reducere", "sum"),
        is_priority=("is_priority", "any")
    ).round(2)

    # Сортировка: сначала приоритетные, по убыванию суммы
    result = result.sort_values(['is_priority', 'Сумма'], ascending=[False, False])
    result_for_save = result.drop(columns=['is_priority'])

    # === ФОРМИРОВАНИЕ КРАСИВОЙ, РОВНОЙ ТАБЛИЦЫ ===
    text = f"📅 Дата отчёта: {report_date}\n\n📊 Отчёт по продажам:\n\n"

    df_display = result_for_save.reset_index()

    # Определяем ширину колонок
    max_name_len = df_display["Denumire marfa"].astype(str).str.len().max()
    name_width = max(max_name_len, len("Denumire marfa")) + 2
    qty_width = 12  # для "Количество"
    sum_width = 10  # для "Сумма"

    # Заголовок
    header = f"{'Denumire marfa':<{name_width}} {'Количество':>{qty_width}} {'Сумма':>{sum_width}}"
    lines = [header, "─" * len(header)]

    # Строки
    for _, row in df_display.iterrows():
        name = str(row["Denumire marfa"])
        qty_str = format_number(row["Количество"])
        sum_str = format_number(row["Сумма"])
        line = f"{name:<{name_width}} {qty_str:>{qty_width}} {sum_str:>{sum_width}}"
        lines.append(line)

    text += "\n".join(lines)
    return report_date, text, result_for_save


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("❌ У вас нет доступа.")
        return
    await update.message.reply_text("Привет! Отправьте .xlsx файл с кассовым отчётом.")


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_authorized(user_id):
        await update.message.reply_text("❌ У вас нет доступа.")
        return

    document = update.message.document
    if not document.file_name.endswith('.xlsx'):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        await update.message.reply_text("📥 Обрабатываю файл...")

        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            await file.download_to_drive(tmp.name)
            input_path = tmp.name

        report_date, text_report, df_result = analyze_excel(input_path)

        # Имя для Excel-файла
        safe_date = "".join(c if c.isalnum() or c in "._-" else "_" for c in report_date)
        output_filename = f"Анализ_отчёта_{safe_date}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        df_result.to_excel(output_path)

        # Отправка текста (если помещается)
        if len(text_report) <= 4090:
            await update.message.reply_text(f"```\n{text_report}\n```", parse_mode="MarkdownV2")
        else:
            await update.message.reply_text("📋 Отчёт слишком длинный для текста. Полная версия — в файле.")

        # Всегда отправляем Excel
        with open(output_path, 'rb') as f:
            await update.message.reply_document(document=f, filename=output_filename)

        # Удалить временные файлы
        os.unlink(input_path)
        os.unlink(output_path)

    except Exception as e:
        logging.exception("Ошибка обработки")
        await update.message.reply_text(f"❌ Ошибка:\n{str(e)[:1000]}")


def main():
    logging.basicConfig(level=logging.INFO)
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_document))

    print("✅ Бот запущен и готов к работе!")
    app.run_polling()


if __name__ == "__main__":
    main()
