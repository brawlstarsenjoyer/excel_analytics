import os
import tempfile
import logging
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# === НАСТРОЙКИ ИЗ ПЕРЕМЕННЫХ ОКРУЖЕНИЯ ===
BOT_TOKEN = os.environ["7109998838:AAGQmN8QyX9wZBI8TPZ0xIvWGHNS8ZA8UXA"]
AUTHORIZED_IDS_STR = os.environ.get("AUTHORIZED_IDS", "")
AUTHORIZED_USER_IDS = set(map(int, AUTHORIZED_IDS_STR.split(","))) if AUTHORIZED_IDS_STR else set()

# Список приоритетных напитков
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


async def analyze_excel(file_path: str) -> tuple[str, str, pd.DataFrame]:
    df_raw = pd.read_excel(file_path, header=None)

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

    report_date = "неизвестна"
    if 'Data' in df.columns:
        non_empty = df['Data'].dropna()
        if not non_empty.empty:
            try:
                report_date = pd.to_datetime(non_empty.iloc[0], dayfirst=True).strftime('%d.%m.%Y')
            except Exception:
                report_date = str(non_empty.iloc[0]).strip()

    required = ["Denumire marfa", "Cantitate", "Suma cu TVA fără reducere"]
    if not all(col in df.columns for col in required):
        raise ValueError("❌ Отсутствуют необходимые столбцы.")

    df = df[required].copy()
    df = df.dropna(subset=["Denumire marfa"])
    df = df[~df["Denumire marfa"].str.contains("Punga", na=False)]

    df['is_priority'] = df['Denumire marfa'].str.lower().str.strip().isin(PRIORITY_DRINKS_LOWER)

    result = df.groupby("Denumire marfa").agg(
        Количество=("Cantitate", "sum"),
        Сумма=("Suma cu TVA fără reducere", "sum"),
        is_priority=("is_priority", "any")
    ).round(2)

    result = result.sort_values(['is_priority', 'Сумма'], ascending=[False, False])
    result_for_save = result.drop(columns=['is_priority'])

    top_rows = result_for_save.head(30)
    text = f"📅 Дата отчёта: {report_date}\n\n📊 Отчёт по продажам:\n\n"
    text += top_rows.to_string()

    if len(result_for_save) > 30:
        text += f"\n\n... и ещё {len(result_for_save) - 30} позиций. Полный отчёт — в файле."

    return report_date, text, result_for_save


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("❌ У вас нет доступа к этому боту.")
        return
    await update.message.reply_text(
        "Привет! Отправьте Excel-файл с кассовым отчётом (.xlsx), и я пришлю анализ."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_authorized(user_id):
        await update.message.reply_text("❌ У вас нет доступа.")
        return

    document = update.message.document
    if not (document.file_name.endswith('.xlsx') or document.mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        await update.message.reply_text("📥 Получаю файл...")

        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            await file.download_to_drive(tmp.name)
            input_path = tmp.name

        report_date, text_report, df_result = analyze_excel(input_path)

        if len(text_report) < 4000:
            await update.message.reply_text(text_report)
        else:
            await update.message.reply_text("Отчёт слишком длинный. Отправляю файл.")

        output_filename = f"Анализ_отчёта_{report_date}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        df_result.to_excel(output_path)

        with open(output_path, 'rb') as f:
            await update.message.reply_document(document=f, filename=output_filename)

        os.unlink(input_path)
        os.unlink(output_path)

    except Exception as e:
        logging.exception("Ошибка обработки")
        await update.message.reply_text(f"❌ Ошибка:\n{str(e)[:1000]}")


def main():
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.MIME_TYPE("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_document))

    print("✅ Бот запущен и ожидает сообщения...")
    app.run_polling()


if __name__ == "__main__":
    main()