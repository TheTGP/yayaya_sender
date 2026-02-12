import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import os
import tempfile
from io import BytesIO
import time

# Настройка страницы
st.set_page_config(
    page_title="📧 Рассылка писем через Яндекс",
    page_icon="📧",
    layout="wide"
)


def create_pdf(text, filename):
    """Создание PDF с текстом"""
    c = canvas.Canvas(filename, pagesize=A4)
    c.setFont("Helvetica", 12)

    text_str = str(text)
    y = 800
    for i in range(0, len(text_str), 90):
        line = text_str[i:i + 90]
        c.drawString(50, y, line)
        y -= 20
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 12)
            y = 800
    c.save()


def send_yandex_email(sender_full, app_password, recipient, text, pdf_path):
    """Отправка письма через Яндекс"""
    login = sender_full.split('@')[0]

    msg = MIMEMultipart()
    msg['From'] = sender_full
    msg['To'] = recipient
    msg['Subject'] = st.session_state.get('email_subject', "Важное сообщение")

    # Текст письма
    body = st.session_state.get('email_body',
                                "Здравствуйте!\n\nВо вложении PDF с вашим сообщением.\n\nС уважением,\nОтправитель")
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    # Прикрепляем PDF
    with open(pdf_path, 'rb') as f:
        attach = MIMEApplication(f.read(), _subtype='pdf')
        attach.add_header('Content-Disposition', 'attachment',
                          filename=f"message_{recipient.split('@')[0]}.pdf")
        msg.attach(attach)

    server = smtplib.SMTP('smtp.yandex.ru', 587)
    server.starttls()
    server.login(login, app_password)
    server.send_message(msg)
    server.quit()


def test_connection(sender_email, app_password):
    """Тестирование подключения к Яндексу"""
    try:
        login = sender_email.split('@')[0]
        server = smtplib.SMTP('smtp.yandex.ru', 587)
        server.starttls()
        server.login(login, app_password)
        server.quit()
        return True, "✅ Подключение успешно!"
    except Exception as e:
        return False, str(e)


# Заголовок
st.title("📧 Массовая рассылка писем через Яндекс")
st.markdown("---")

# Сайдбар для настроек
with st.sidebar:
    st.header("⚙️ Настройки отправителя")

    sender_email = st.text_input(
        "📧 Ваш Яндекс email",
        placeholder="your.email@yandex.ru",
        help="Полный адрес Яндекс почты"
    )

    app_password = st.text_input(
        "🔑 Пароль приложения",
        type="password",
        placeholder="Скопируйте пароль из Яндекса",
        help="Не ваш обычный пароль, а специальный пароль приложения!"
    )

    st.markdown("---")
    st.header("📝 Настройки письма")

    email_subject = st.text_input(
        "✉️ Тема письма",
        value="Важное сообщение",
        key="email_subject"
    )

    email_body = st.text_area(
        "📄 Текст письма",
        value="Здравствуйте!\n\nВо вложении PDF с вашим сообщением.\n\nС уважением,\nОтправитель",
        height=150,
        key="email_body"
    )

    st.markdown("---")
    st.header("ℹ️ Инструкция")
    st.info("""
    1. Включите доступ в настройках Яндекса
    2. Создайте пароль приложения
    3. Загрузите Excel файл
    4. Нажмите "Начать рассылку"

    **Формат Excel:**
    - Колонка A: Email получателя
    - Колонка B: Текст для PDF
    """)

# Основной контент
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📤 Загрузите файл с получателями")

    uploaded_file = st.file_uploader(
        "Выберите Excel файл",
        type=['xlsx', 'xls'],
        help="Файл должен содержать email в колонке A и текст в колонке B"
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, header=None)
            st.success(f"✅ Загружено {len(df)} получателей")

            # Показываем preview
            st.subheader("📋 Предпросмотр данных")
            preview_df = df.iloc[:5].copy()
            preview_df.columns = ['Email', 'Текст сообщения'] if len(df.columns) > 1 else ['Email', 'Нет данных']
            st.dataframe(preview_df, use_container_width=True)

            if len(df.columns) < 2:
                st.warning("⚠️ В файле только одна колонка. Убедитесь, что текст для PDF есть в колонке B")

            # Сохраняем в session state
            st.session_state['df'] = df

        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")

with col2:
    st.header("🔄 Проверка подключения")

    if st.button("🔌 Проверить подключение к Яндекс", use_container_width=True):
        if not sender_email or not app_password:
            st.error("❌ Заполните email и пароль приложения!")
        else:
            with st.spinner("Проверяем подключение..."):
                success, message = test_connection(sender_email, app_password)
                if success:
                    st.success(message)
                else:
                    st.error(f"❌ Ошибка: {message}")
                    st.info("""
                    💡 Возможные причины:
                    1. Неправильный пароль приложения
                    2. Не включен доступ в настройках Яндекса
                    3. Логин должен быть без @yandex.ru
                    """)

# Кнопка отправки
st.markdown("---")

if 'df' in st.session_state:
    col_send1, col_send2, col_send3 = st.columns([1, 2, 1])
    with col_send2:
        if st.button("🚀 НАЧАТЬ МАССОВУЮ РАССЫЛКУ",
                     type="primary",
                     use_container_width=True,
                     disabled=not (sender_email and app_password)):

            df = st.session_state['df']

            # Прогресс бар
            progress_bar = st.progress(0, text="Подготовка к отправке...")
            status_text = st.empty()

            # Создаем временную папку
            with tempfile.TemporaryDirectory() as temp_folder:
                success_count = 0
                fail_count = 0
                results = []

                # Результаты в expander
                with st.expander("📨 Детали отправки", expanded=True):
                    results_container = st.container()

                for index, row in df.iterrows():
                    recipient = str(row.iloc[0]).strip()
                    text = row.iloc[1] if len(row) > 1 else ""

                    # Пропускаем пустые
                    if not recipient or recipient == '' or pd.isna(recipient):
                        results.append({"email": "Пустой", "status": "⚠️ Пропущен", "error": "Пустой email"})
                        continue

                    # Обновляем прогресс
                    progress = (index + 1) / len(df)
                    progress_bar.progress(progress, text=f"Отправка {index + 1}/{len(df)}: {recipient}")
                    status_text.text(f"📨 Отправляем письмо {index + 1} из {len(df)}...")

                    try:
                        # Создаем PDF
                        pdf_name = f"temp_{index}.pdf"
                        pdf_path = os.path.join(temp_folder, pdf_name)
                        create_pdf(text, pdf_path)

                        # Отправляем
                        send_yandex_email(sender_email, app_password, recipient, text, pdf_path)

                        success_count += 1
                        results.append({"email": recipient, "status": "✅ Успешно", "error": ""})

                    except Exception as e:
                        fail_count += 1
                        results.append({"email": recipient, "status": "❌ Ошибка", "error": str(e)})

                    # Показываем последние результаты
                    with results_container:
                        results_df = pd.DataFrame(results[-10:])  # последние 10
                        st.dataframe(results_df, use_container_width=True)

                # Финальный результат
                progress_bar.progress(1.0, text="✅ Рассылка завершена!")

                st.markdown("---")
                st.subheader("📊 Итоговые результаты")

                col_res1, col_res2, col_res3 = st.columns(3)
                with col_res1:
                    st.metric("✅ Успешно отправлено", success_count)
                with col_res2:
                    st.metric("❌ Ошибок", fail_count)
                with col_res3:
                    st.metric("📧 Всего обработано", len(df))

                # Скачать отчет
                if results:
                    results_full_df = pd.DataFrame(results)
                    csv = results_full_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Скачать отчет о рассылке",
                        data=csv,
                        file_name="report_sending.csv",
                        mime="text/csv"
                    )
else:
    st.info("👆 Загрузите Excel файл, чтобы начать рассылку")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        ⚡ Разработано для массовой рассылки через Яндекс почту<br>
        📌 Не забудьте создать <b>пароль приложения</b> в настройках Яндекса!
    </div>
    """,
    unsafe_allow_html=True
)