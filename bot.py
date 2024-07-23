import telebot
import re
import pandas as pd
bot = telebot.TeleBot("6626124338:AAGXzVqsvYqjOq0PeRkipueDbc7UViQ15Bk")


questions_file_path = 'questions.xlsx'

df = pd.read_excel(questions_file_path)

ADMIN_CHAT_ID = 668290718
keywords_responses = list(df.itertuples(index=False, name=None))

def reload_questions():
    global questions_data
    questions_data = pd.read_excel(questions_file_path)


@bot.message_handler(commands=['admin'])
def handle_admin(message):
    if message.chat.id == ADMIN_CHAT_ID:
        bot.send_document(message.chat.id, open(questions_file_path, 'rb'))
        bot.send_message(message.chat.id, "Iltimos, faylni o'zgartiring va savollarni yangilash uchun uni qayta yuklang.")
    else:
        bot.send_message(message.chat.id, "Siz ushbu buyruqqa kirish huquqiga ega emassiz.")


@bot.message_handler(content_types=['document'])
def handle_document(message):
    if message.chat.id == ADMIN_CHAT_ID:
        try:
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)

            with open(questions_file_path, 'wb') as new_file:
                new_file.write(downloaded_file)

            reload_questions()
            bot.send_message(message.chat.id, "Savollar muvaffaqiyatli yangilandi.")
        except Exception as e:
            bot.send_message(message.chat.id, f"Xatolik yuz berdi: {e}")
    else:
        return None



@bot.message_handler(func=lambda message: True, content_types=['text'])
def handle_text(message):
    response = generate_response(message.text)
    if response:
        bot.reply_to(message, response)


def generate_response(question):
    for keywords, response in keywords_responses:
        keyword_list = [keyword.strip() for keyword in keywords.split(',')]
        if all(re.search(fr'\b{keyword}\b', question, re.IGNORECASE) for keyword in keyword_list):
            return response
    return None




bot.polling(none_stop=True)




