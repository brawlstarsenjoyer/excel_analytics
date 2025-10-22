import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes

# 🔑 Вставь сюда свой токен от @BotFather
BOT_TOKEN = "7109998838:AAGQmN8QyX9wZBI8TPZ0xIvWGHNS8ZA8UXA"

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    file = update.message.document

    if not file.file_name.endswith('.xlsx'):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        # Скачиваем файл
        new_file = await context.bot.get_file(file.file_id)
        file_path = f"{user.id}_report.xlsx"
        await new_file.download_to_drive(file_path)

        # Читаем Excel
        df = pd.read_excel(file_path, header=None)

        # Определяем заголовки (в вашем файле данные начинаются со 2-й строки)
        # Ищем строку с "Denumire marfa"
        header_row = None
        for i, row in df.iterrows():
            if "Denumire marfa" in row.values:
                header_row = i
                break

        if header_row is None:
            await update.message.reply_text("Не удалось найти заголовки в файле.")
            os.remove(file_path)
            return

        df.columns = df.iloc[header_row]
        df = df[header_row + 1:].reset_index(drop=True)

        # Убираем строки без названия товара
        df = df.dropna(subset=["Denumire marfa"])

        # Исключаем "Punga", если нужно
        df = df[~df["Denumire marfa"].str.contains("Punga", na=False)]

        # Агрегация
        result = df.groupby("Denumire marfa").agg(
            Количество=("Cantitate", "sum"),
            Сумма=("Suma cu TVA fără reducere", "sum")
        ).round(2).sort_values("Сумма", ascending=False)

        # Формируем текст
        text = "📊 Отчёт по продажам:\n\n"
        for idx, row in result.iterrows():
            text += f"• {idx}\n  Кол-во: {int(row['Количество'])}, Сумма: {row['Сумма']:.0f}\n\n"

        if len(text) > 4096:
            text = text[:4090] + "\n[... обрезано]"

        await update.message.reply_text(text)

    except Exception as e:
        await update.message.reply_text(f"Ошибка при обработке файла: {str(e)}")
    finally:
        # Удаляем временный файл
        if os.path.exists(file_path):
            os.remove(file_path)

# Запуск бота
if __name__ == "__main__":
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.MIME_TYPE("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))
    print("Бот запущен...")
    app.run_polling()