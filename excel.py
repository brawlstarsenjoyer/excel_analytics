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

# === Список приоритетных напитков (ваши) ===
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
    """Проверяет, разрешён ли пользователь."""
    return not AUTHORIZED_USER_IDS or user_id in AUTHORIZED_USER_IDS


def analyze_excel(file_path: str) -> tuple[str, str, pd.DataFrame]:
    """
    Анализирует Excel-файл и возвращает:
    - дату отчёта (str)
    - текстовый отчёт (str)
    - датафрейм для сохранения (pd.DataFrame)
    """
    df_raw = pd.read_excel(file_path, header=None)

    # Найти строку с заголовками
    header_row = None
    for i in range(len(df_raw)):
        if "Denumire marfa" in df_raw.iloc[i].values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("❌ Не найдены заголовки. Убедитесь, что файл — отчёт кассы.")

    # Установить заголовки
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

    # Определить приоритетные напитки
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

    # Текстовый отчёт (макс. 30 строк)
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
    if not document.file_name.endswith('.xlsx'):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        await update.message.reply_text("📥 Получаю и обрабатываю файл...")

        # Скачать файл
        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            await file.download_to_drive(tmp.name)
            input_path = tmp.name

        # Анализ (синхронно!)
        report_date, text_report, df_result = analyze_excel(input_path)

        # Подготовить имя файла
        safe_date = report_date.replace("/", "-").replace(":", "-")
        output_filename = f"Анализ_отчёта_{safe_date}.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        df_result.to_excel(output_path)

        # Отправить текст (если помещается)
        if len(text_report) < 4000:
            await update.message.reply_text(text_report)
        else:
            await update.message.reply_text("Отчёт слишком длинный для текста. Смотрите Excel-файл.")

        # Отправить Excel
        with open(output_path, 'rb') as f:
            await update.message.reply_document(document=f, filename=output_filename)

        # Удалить временные файлы
        os.unlink(input_path)
        os.unlink(output_path)

    except Exception as e:
        logging.exception("Ошибка при обработке файла")
        await update.message.reply_text(f"❌ Ошибка обработки:\n{str(e)[:1000]}")


def main():
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )

    app = Application.builder().token(BOT_TOKEN).build()

    # Обработчики
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_document))

    print("✅ Бот запущен и ожидает файлы...")
    app.run_polling()


if __name__ == "__main__":
    main()
