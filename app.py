
import streamlit as st
import pandas as pd
import json
import os
import base64
import re

from io import BytesIO
from google import genai
from pptx import Presentation

# ==================================
# 1. APP SOZLAMALARI
# ==================================

st.set_page_config(
    page_title="JON AI 3.5",
    page_icon="💎",
    layout="wide"
)

APP_VERSION = "3.5"
MODEL = "gemini-2.5-flash"

JON_CORE = """
Sen JON AI 3.5 nomli aqlli virtual yordamchisan.
O'zbek tilida ravon gapirasan.
Savollarga aniq va tushunarli javob berasan.
Ta'lim, tarjima, dasturlash va ijodiy ishlarda yordam berasan.
Doimo hurmat bilan muloqot qilasan.
"""

# ==================================
# 2. SESSION STATE
# ==================================

defaults = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "avatar": None,
    "api_key": ""
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ==================================
# 3. GEMINI AI
# ==================================

def ask_ai(prompt, extra_context=""):

    api_key = st.session_state.get(
        "api_key", ""
    ).strip()

    if not api_key:
        return "API kalitini Sozlamalardan kiriting."

    try:
        client = genai.Client(api_key=api_key)

        response = client.models.generate_content(
            model=MODEL,
            contents=(
                JON_CORE
                + "\n"
                + extra_context
                + "\n"
                + prompt
            )
        )

        if response.text:
            return response.text

        return "AI javob qaytarmadi."

    except Exception as e:
        return f"Gemini xatoligi: {str(e)}"


# ==================================
# 4. PERSONAJ
# ==================================

def show_character(answer=""):

    avatar = st.session_state.get("avatar")

    if avatar:
        avatar_data = base64.b64encode(
            avatar
        ).decode("utf-8")

        avatar_html = f"""
        <img src="data:image/png;base64,{avatar_data}"
        style="
        width:280px;
        height:300px;
        object-fit:contain;
        filter:drop-shadow(0 0 25px #00aaff);
        ">
        """
    else:
        avatar_html = """
        <div style="
        font-size:150px;
        padding:30px;
        filter:drop-shadow(0 0 25px #008cff);
        ">🤖</div>
        """

    name = st.session_state.character_name

    st.markdown(
        f"""
        <div style="
        background:radial-gradient(circle,#10285c,#030817);
        border:1px solid #008cff;
        border-radius:20px;
        padding:20px;
        text-align:center;
        color:white;
        ">
        {avatar_html}
        <h2 style="color:#00dfff">{name} AI</h2>
        <p style="color:#00eaff">
        ● ONLINE
        </p>
        <p>{st.session_state.character_style}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if answer:
        st.info(answer)

        if st.button("🔊 Javobni eshitish"):
            st.components.v1.html(
                f"""
                <script>
                const speech = new SpeechSynthesisUtterance(
                    {json.dumps(answer, ensure_ascii=False)}
                );
                speech.lang = "uz-UZ";
                speechSynthesis.cancel();
                speechSynthesis.speak(speech);
                </script>
                """,
                height=0
            )


# ==================================
# 5. POWERPOINT JSON PARSER
# ==================================

def parse_slides(result):

    try:
        clean = result.strip()

        clean = re.sub(
            r"```(?:json)?",
            "",
            clean
        )

        clean = clean.replace("```", "").strip()

        start = clean.find("[")
        end = clean.rfind("]") + 1

        if start == -1 or end == 0:
            raise ValueError("JSON massiv topilmadi")

        data = json.loads(clean[start:end])

        if not isinstance(data, list):
            raise ValueError("Natija ro'yxat emas")

        return data

    except Exception as e:
        st.error(f"JSON xatosi: {e}")
        return []


# ==================================
# 6. POWERPOINT YARATISH
# ==================================

def create_presentation(topic, slides_data):

    prs = Presentation()

    for item in slides_data:

        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )

        slide.shapes.title.text = str(
            item.get("title", topic)
        )

        body = slide.placeholders[1].text_frame

        points = item.get("points", [])

        if isinstance(points, str):
            points = [points]

        body.text = ""

        for point in points:

            paragraph = body.add_paragraph()
            paragraph.text = str(point)
            paragraph.level = 0

    output = BytesIO()

    prs.save(output)
    output.seek(0)

    return output


# ==================================
# 7. SIDEBAR
# ==================================

pages = [
    "🏠 Bosh sahifa",
    "🤖 AI Chat",
    "🎓 Ta'lim markazi",
    "📊 PowerPoint",
    "🌐 Tarjimon",
    "🎨 Rasm prompti",
    "📈 Excel tahlil",
    "📝 Vazifalar",
    "📷 Kamera",
    "👤 Personaj",
    "⚙️ Sozlamalar",
    "💎 Premium"
]

st.sidebar.title("💎 JON AI 3.5")
st.sidebar.caption("Virtual Intelligence System")

page = st.sidebar.radio(
    "Bo'limni tanlang",
    pages
)

st.sidebar.divider()

st.sidebar.caption(
    "JON AI 3.5 | 2026"
)


# ==================================
# 8. BOSH SAHIFA
# ==================================

if page == "🏠 Bosh sahifa":

    st.title("💎 JON AI 3.5")

    st.caption(
        "Sizning shaxsiy virtual yordamchingiz"
    )

    show_character()

    st.markdown("### Xush kelibsiz!")

    st.write(
        "JON AI bilan suhbatlashing, bilim oling, "
        "tarjima qiling va yangi loyihalar yarating."
    )

    st.success("JON AI tizimi tayyor!")


# ==================================
# 9. AI CHAT
# ==================================

elif page == "🤖 AI Chat":

    st.title("🤖 JON AI — Aqlli suhbat")

    for message in st.session_state.messages:

        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    prompt = st.chat_input(
        "JON AI ga xabar yozing..."
    )

    if prompt:

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):

            with st.spinner("JON AI javob tayyorlamoqda..."):

                answer = ask_ai(prompt)

                st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })

        st.rerun()


# ==================================
# 10. TA'LIM MARKAZI
# ==================================

elif page == "🎓 Ta'lim markazi":

    st.title("🎓 JON AI — Ta'lim markazi")

    subject = st.selectbox(
        "Fan yoki yo'nalish",
        [
            "Matematika",
            "Ingliz tili",
            "Fizika",
            "Kimyo",
            "Tarix",
            "Informatika",
            "Psixologiya",
            "Boshqa"
        ]
    )

    question = st.text_area(
        "Savolingiz yoki mavzuni yozing"
    )

    if st.button("📚 Tushuntirish"):

        if question:

            with st.spinner("Tayyorlanmoqda..."):

                answer = ask_ai(
                    question,
                    "Fan: " + subject +
                    ". Mavzuni o'quvchiga sodda tushuntir."
                )

                st.markdown(answer)

        else:
            st.warning("Avval savol yozing.")


# ==================================
# 11. POWERPOINT
# ==================================

elif page == "📊 PowerPoint":

    st.title("📊 JON AI — PowerPoint yaratuvchi")

    topic = st.text_input(
        "Taqdimot mavzusi"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        3, 15, 6
    )

    if st.button("✨ Taqdimot yaratish"):

        if topic:

            prompt = (
                "Mavzu: " + topic +
                ". " + str(slide_count) +
                " ta slayd tayyorla. "
                "Faqat JSON massiv qaytar. "
                'Har bir elementda "title" va "points" bo\'lsin. '
                "points ro'yxat ko'rinishida bo'lsin."
            )

            with st.spinner("Taqdimot tayyorlanmoqda..."):

                result = ask_ai(prompt)

                slides_data = parse_slides(result)

                if slides_data:

                    pptx_file = create_presentation(
                        topic,
                        slides_data
                    )

                    st.success("Taqdimot tayyor!")

                    st.download_button(
                        "⬇️ PowerPoint yuklab olish",
                        data=pptx_file,
                        file_name="JON_AI_Taqdimot.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

        else:
            st.warning("Mavzuni kiriting.")


# ==================================
# 12. TARJIMON
# ==================================

elif page == "🌐 Tarjimon":

    st.title("🌐 JON AI — Universal tarjimon")

    source_lang = st.selectbox(
        "Matn qaysi tilda?",
        [
            "Avtomatik aniqlash",
            "O'zbek",
            "English",
            "Русский",
            "Turk",
            "Arab",
            "Nemis",
            "Fransuz"
        ]
    )

    target_lang = st.selectbox(
        "Qaysi tilga tarjima qilinsin?",
        [
            "O'zbek",
            "English",
            "Русский",
            "Turk",
            "Arab",
            "Nemis",
            "Fransuz"
        ]
    )

    text_to_translate = st.text_area(
        "Tarjima qilinadigan matn"
    )

    if st.button("🌐 Tarjima qilish"):

        if text_to_translate:

            prompt = (
                "Quyidagi matnni "
                + source_lang
                + " tilidan "
                + target_lang
                + " tiliga tarjima qil. "
                "Faqat tarjimani chiqar:\n"
                + text_to_translate
            )

            with st.spinner("Tarjima qilinmoqda..."):

                translated = ask_ai(prompt)

                st.markdown("### Natija")
                st.write(translated)

        else:
            st.warning("Matn kiriting.")


# ==================================
# 13. RASM PROMPTI
# ==================================

elif page == "🎨 Rasm prompti":

    st.title("🎨 JON AI — Rasm prompti")

    idea = st.text_area(
        "Qanday rasm yaratmoqchisiz?"
    )

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Realistik",
            "3D",
            "Multfilm",
            "Anime",
            "Kiberpank",
            "Fantastik"
        ]
    )

    if st.button("✨ Prompt yaratish"):

        if idea:

            prompt = (
                "Quyidagi g'oya uchun batafsil "
                "rasm prompti yoz. Uslub: "
                + style
                + ". G'oya: "
                + idea
            )

            with st.spinner("Prompt yaratilmoqda..."):

                result = ask_ai(prompt)

                st.markdown("### Tayyor prompt")
                st.code(result)

                st.button(
                    "📋 Promptni nusxalash",
                    on_click=st.toast,
                    args=("Prompt tayyor!",)
                )

        else:
            st.warning("Rasm g'oyasini yozing.")


# ==================================
# 14. EXCEL TAHLIL
# ==================================

elif page == "📈 Excel tahlil":

    st.title("📈 JON AI — Excel tahlil")

    uploaded_file = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "csv"]
    )

    if uploaded_file:

        try:

            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("Fayl muvaffaqiyatli ochildi!")

            st.dataframe(
                df,
                use_container_width=True
            )

            st.markdown("### Asosiy ma'lumotlar")

            st.write(
                df.describe(include="all")
            )

            question = st.text_input(
                "Jadval haqida savolingiz"
            )

            if st.button("📊 Tahlil qilish"):

                if question:

                    context = df.head(30).to_string()

                    answer = ask_ai(
                        question,
                        "Jadval ma'lumotlari:\n" + context
                    )

                    st.markdown(answer)

                else:
                    st.warning("Savol kiriting.")

        except Exception as e:
            st.error(f"Fayl xatosi: {e}")


# ==================================
# 15. VAZIFALAR
# ==================================

elif page == "📝 Vazifalar":

    st.title("📝 JON AI — Kunlik vazifalar")

    new_task = st.text_input(
        "Yangi vazifa"
    )

    if st.button("➕ Vazifa qo'shish"):

        if new_task:

            st.session_state.tasks.append(new_task)
            st.rerun()

    st.markdown("### Vazifalar ro'yxati")

    for i, task in enumerate(
        st.session_state.tasks
    ):

        st.checkbox(
            task,
            key=f"task_{i}"
        )

    if st.button("🗑️ Barcha vazifalarni tozalash"):

        st.session_state.tasks = []
        st.rerun()


# ==================================
# 16. KAMERA
# ==================================

elif page == "📷 Kamera":

    st.title("📷 JON AI — Kamera")

    photo = st.camera_input(
        "Suratga olish"
    )

    if photo:

        st.image(
            photo,
            caption="Olingan surat",
            use_container_width=True
        )


# ==================================
# 17. PERSONAJ
# ==================================

elif page == "👤 Personaj":

    st.title("👤 JON AI — Personaj sozlamalari")

    st.text_input(
        "Personaj nomi",
        key="character_name"
    )

    st.selectbox(
        "Tashqi ko'rinish",
        [
            "Hologramma",
            "Robot",
            "Multfilm uslubi",
            "Fantastik"
        ],
        key="character_style"
    )

    avatar = st.file_uploader(
        "Personaj rasmi yuklash",
        type=["png", "jpg", "jpeg"]
    )

    if avatar:
        st.session_state.avatar = avatar.getvalue()

    if st.session_state.avatar:

        st.image(
            st.session_state.avatar,
            caption="Siz tanlagan personaj",
            use_container_width=True
        )

    else:
        show_character()

    st.info(
        "Personajning nomi, uslubi va rasmi "
        "shu bo'limda sozlanadi."
    )


# ==================================
# 18. SOZLAMALAR
# ==================================

elif page == "⚙️ Sozlamalar":

    st.title("⚙️ JON AI — Sozlamalar")

    st.markdown("### Gemini API")

    api_input = st.text_input(
        "Gemini API kaliti",
        type="password",
        value=st.session_state.api_key
    )

    if st.button("💾 API kalitini saqlash"):

        st.session_state.api_key = api_input.strip()

        st.success("API kaliti saqlandi!")

    if st.button("🔄 API ulanishini tekshirish"):

        if st.session_state.api_key:

            result = ask_ai(
                "Salom. Faqat OK deb javob ber."
            )

            if result.startswith("Gemini xatoligi:"):
                st.error(result)
            elif result == "API kalitini Sozlamalardan kiriting.":
                st.warning(result)
            else:
                st.success("API javob berdi!")
                st.write(result)

        else:
            st.warning("API kalitini kiriting!")

    st.divider()

    st.write("Ilova versiyasi:", APP_VERSION)
    st.write("Model:", MODEL)


# ==================================
# 19. PREMIUM
# ==================================

elif page == "💎 Premium":

    st.title("💎 JON AI Premium")

    st.subheader("Kelajakdagi imkoniyatlar")

    st.write("🚀 Kengaytirilgan AI imkoniyatlari")
    st.write("🎭 Maxsus personaj sozlamalari")
    st.write("🎙️ Ovozli yordamchi")
    st.write("✨ Maxsus dizayn va effektlar")
    st.write("📊 Kengaytirilgan tahlil")

    st.info(
        "Premium bo'limi hozircha namoyish uchun. "
        "To'lov tizimi hali ulanmagan."
    )


# ==================================
# 20. FOOTER
# ==================================

st.sidebar.divider()

st.sidebar.caption(
    "💎 JON AI — Your Virtual Assistant"
)

st.sidebar.caption(
    "Powered by Google Gemini"
)
