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
    page_title="JON AI LIVE",
    page_icon="💎",
    layout="wide"
)

APP_VERSION = "5.0"
MODEL = "gemini-3.5-flash-lite"

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

def ask_ai(prompt, extra_context="", history=None):

    try:
        saved_secret = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        saved_secret = ""

    api_key = (
        st.session_state.get("api_key", "") or saved_secret
    ).strip()

    if not api_key:
        return "API kalitini Sozlamalardan kiriting."

    try:
        client = genai.Client(api_key=api_key)

        conversation = []
        for item in (history or [])[-12:]:
            role = "Foydalanuvchi" if item.get("role") == "user" else "JON AI"
            conversation.append(f"{role}: {item.get('content', '')}")

        contents = (
            JON_CORE
            + "\n"
            + ("\n".join(conversation) + "\n" if conversation else "")
            + extra_context
            + "\nFoydalanuvchi: "
            + prompt
        )

        response = client.models.generate_content(
            model=MODEL,
            contents=contents
        )

        if response.text:
            return response.text

        return "AI javob qaytarmadi."

    except Exception as e:
        return f"Gemini xatoligi: {str(e)}"


# ==================================
# 4. PERSONAJ
# ==================================
        stream = null;
        const video = document.getElementById('camera');
        video.srcObject = null;
        video.style.display = 'none';
        document.getElementById('camera-note').textContent = 'Kamera o‘chirildi.';
      };
      if(speakText) welcome.textContent = speakText;
      if(!speakText) {
        root.classList.add('dancing');
        setTimeout(()=>root.classList.remove('dancing'),5000);
      }
      let blinkTimer=setInterval(()=> {
        root.classList.add('blink');
        setTimeout(()=>root.classList.remove('blink'),160);
      },4200);
      window.addEventListener('beforeunload',()=>clearInterval(blinkTimer));
    })();
    </script>
    """

    html = (html.replace("__COLOR__", color)
                .replace("__NAME__", name)
                .replace("__ANSWER__", json.dumps(str(spoken_text), ensure_ascii=False)))

    components.html(html, height=760, scrolling=False)

    if answer:
        st.info(answer)


def parse_slides(result):
    try:
        clean = result.strip()
        clean = re.sub(r"```(?:json)?", "", clean)
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
        slide = prs.slides.add_slide(prs.slide_layouts[1])

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
          document.getElementById('camera-off').onclick=()=> {
        if(stream) stream.getTracks().forEach(t=>t.stop());
        stream = null;
        const video = document.getElementById('camera');
        video.srcObject = null;
        video.style.display = 'none';
        document.getElementById('camera-note').textContent = 'Kamera o‘chirildi.';
      };
      if(speakText) welcome.textContent = speakText;
      if(!speakText) {
        root.classList.add('dancing');
        setTimeout(()=>root.classList.remove('dancing'),5000);
      }
      let blinkTimer=setInterval(()=> {
        root.classList.add('blink');
        setTimeout(()=>root.classList.remove('blink'),160);
      },4200);
      window.addEventListener('beforeunload',()=>clearInterval(blinkTimer));
    })();
    </script>
    """

    html = (html.replace("__COLOR__", color)
                .replace("__NAME__", name)
                .replace("__ANSWER__", json.dumps(str(spoken_text), ensure_ascii=False)))

    components.html(html, height=760, scrolling=False)

    if answer:
        st.info(answer)


def parse_slides(result):
    try:
        clean = result.strip()
        clean = re.sub(r"```(?:json)?", "", clean)
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


def create_presentation(topic, slides_data):
    prs = Presentation()

    for item in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

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

st.sidebar.title("💎 JON AI 5.0")
st.sidebar.caption("Virtual Intelligence System")

page = st.sidebar.radio(
    "Bo'limni tanlang",
    pages
)

# ==================================
# 8. BOSH SAHIFA
# ==================================

if page == "🏠 Bosh sahifa":

    st.title("💎 JON AI LIVE")
    st.caption("Sizning shaxsiy virtual yordamchingiz")

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

    prompt = st.chat_input("JON AI ga xabar yozing...")

    if prompt:
        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI javob tayyorlamoqda..."):
                answer = ask_ai(
                    prompt,
                    history=st.session_state.messages[:-1]
                )
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
                    extra_context=f"Fan: {subject}"
                )
            st.markdown(answer)
        else:
            st.warning("Avval savolingizni yozing.")
            # ==================================
# 11. POWERPOINT
# ==================================

elif page == "📊 PowerPoint":

    st.title("📊 JON AI — PowerPoint yaratish")

    topic = st.text_input(
        "Prezentatsiya mavzusi"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        3, 15, 7
    )

    if st.button("✨ Prezentatsiya yaratish"):

        if topic:
            with st.spinner("Slaydlar tayyorlanmoqda..."):

                prompt = f"""
                {topic} mavzusida {slide_count} ta slayd
                uchun JSON formatida ma'lumot yarat.

                Har bir elementda:
                title va points bo'lsin.
                Faqat JSON massiv qaytar.
                """

                result = ask_ai(prompt)
                slides = parse_slides(result)

                if slides:
                    pptx_file = create_presentation(
                        topic,
                        slides
                    )

                    st.download_button(
                        "📥 PowerPoint yuklab olish",
                        data=pptx_file,
                        file_name="JON_AI_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
        else:
            st.warning("Mavzuni kiriting.")


# ==================================
# 12. TARJIMON
# ==================================

elif page == "🌐 Tarjimon":

    st.title("🌐 JON AI — Tarjimon")

    source_language = st.selectbox(
        "Qaysi tildan?",
        ["Avtomatik aniqlash", "O'zbek", "Ingliz", "Rus", "Turk"]
    )

    target_language = st.selectbox(
        "Qaysi tilga?",
        ["O'zbek", "Ingliz", "Rus", "Turk"]
    )

    text_input = st.text_area(
        "Tarjima qilinadigan matn"
    )

    if st.button("🌍 Tarjima qilish"):

        if text_input:
            with st.spinner("Tarjima qilinmoqda..."):

                translation = ask_ai(
                    text_input,
                    extra_context=f"""
                    Matnni {source_language} tilidan
                    {target_language} tiliga tarjima qil.
                    Faqat tarjimani qaytar.
                    """
                )

            st.success("Tarjima tayyor!")
            st.write(translation)
        else:
            st.warning("Matn kiriting.")


# ==================================
# 13. RASM PROMPTI
# ==================================

elif page == "🎨 Rasm prompti":

    st.title("🎨 AI rasm g'oyasi")

    idea = st.text_area(
        "Qanday rasm yaratmoqchisiz?"
    )

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Realistik",
            "3D",
            "Anime",
            "Kinematik",
            "Fantastik",
            "Minimalistik"
        ]
    )

    if st.button("🎨 Prompt yaratish"):

        if idea:
            with st.spinner("Prompt tayyorlanmoqda..."):

                result = ask_ai(
                    idea,
                    extra_context=f"""
                    Ushbu g'oya uchun batafsil
                    AI image generation prompt yoz.
                    Uslub: {style}.
                    """
                )

            st.markdown("### Tayyor prompt")
            st.write(result)

            st.button(
                "📋 Nusxalash",
                on_click=lambda: st.session_state.update(
                    {"copied_prompt": result}
                )
            )

            st.caption("Promptni rasm yaratish vositasiga kiritishingiz mumkin.")

        else:
            st.warning("Rasm g'oyasini yozing.")
            # ==================================
# 14. EXCEL TAHLIL
# ==================================

elif page == "📈 Excel tahlil":

    st.title("📈 JON AI — Excel tahlil")

    uploaded_file = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file:

        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.subheader("Ma'lumotlar")
            st.dataframe(df, use_container_width=True)

            st.subheader("Umumiy ma'lumot")
            st.write(df.describe(include="all"))

            if st.button("🤖 AI orqali tahlil"):

                with st.spinner("Tahlil qilinmoqda..."):

                    summary = df.head(30).to_string()

                    result = ask_ai(
                        "Ushbu jadvalni tahlil qil.",
                        extra_context=summary
                    )

                st.markdown(result)

        except Exception as e:
            st.error(f"Faylni o'qishda xatolik: {e}")


# ==================================
# 15. VAZIFALAR
# ==================================

elif page == "📝 Vazifalar":

    st.title("📝 JON AI — Vazifalar")

    task = st.text_input(
        "Yangi vazifa"
    )

    if st.button("➕ Qo'shish"):

        if task:
            st.session_state.tasks.append({
                "name": task,
                "done": False
            })

            st.rerun()

    st.subheader("Vazifalar ro'yxati")

    for i, item in enumerate(st.session_state.tasks):

        col1, col2 = st.columns([4, 1])

        with col1:
            checked = st.checkbox(
                item["name"],
                value=item["done"],
                key=f"task_{i}"
            )

            st.session_state.tasks[i]["done"] = checked

        with col2:
            if st.button("🗑️", key=f"delete_{i}"):

                st.session_state.tasks.pop(i)
                st.rerun()


# ==================================
# 16. KAMERA
# ==================================

elif page == "📷 Kamera":

    st.title("📷 JON AI — Kamera")

    picture = st.camera_input(
        "Suratga olish"
    )

    if picture:

        st.image(
            picture,
            caption="Yuklangan surat",
            use_container_width=True
        )

        st.success("Surat muvaffaqiyatli olindi!")

        if st.button("🤖 Surat haqida so'rash"):

            st.info(
                "Hozirgi versiyada suratni AI orqali "
                "tahlil qilish funksiyasi ulanmagan."
            )
            # ==================================
# 17. PERSONAJ
# ==================================

elif page == "👤 Personaj":

    st.title("👤 JON AI — Personaj sozlamalari")

    st.session_state.character_name = st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    st.session_state.character_style = st.selectbox(
        "Personaj ko'rinishi",
        [
            "Hologramma",
            "Robot",
            "Multfilm uslubi",
            "Fantastik"
        ]
    )

    avatar_file = st.file_uploader(
        "Personaj suratini yuklash",
        type=["png", "jpg", "jpeg"]
    )

    if avatar_file:
        st.session_state.avatar = avatar_file.getvalue()
        st.image(st.session_state.avatar, width=200)

    if st.button("💾 Saqlash"):
        st.success("Personaj sozlamalari saqlandi!")


# ==================================
# 18. SOZLAMALAR
# ==================================

elif page == "⚙️ Sozlamalar":

    st.title("⚙️ JON AI — Sozlamalar")

    api_key = st.text_input(
        "Gemini API kaliti",
        value=st.session_state.api_key,
        type="password"
    )

    if st.button("💾 API kalitini saqlash"):

        st.session_state.api_key = api_key.strip()
        st.success("API kaliti joriy sessiyada saqlandi!")

    if st.button("🧪 AI ni tekshirish"):

        with st.spinner("Tekshirilmoqda..."):

            result = ask_ai(
                "Salom! O'zingni qisqacha tanishtir."
            )

        st.write(result)


# ==================================
# 19. PREMIUM
# ==================================

elif page == "💎 Premium":

    st.title("💎 JON AI PREMIUM")

    st.subheader("Premium imkoniyatlar")

    st.write("""
    - Kengaytirilgan AI suhbat
    - Professional ta'lim yordami
    - PowerPoint yaratish
    - Excel tahlil
    - Tarjima xizmatlari
    - Kelajakda ovozli AI imkoniyatlari
    """)

    st.warning(
        "Premium to'lov tizimi hali ulanmagan."
    )

    st.info(
        "Bu bo'lim hozircha ma'lumot uchun."
    )


# ==================================
# 20. YAKUNIY MA'LUMOT
# ==================================

st.sidebar.divider()

st.sidebar.caption(
    f"JON AI VERSION {APP_VERSION}"
)

st.sidebar.caption(
    "Created with Python and Streamlit"
    )
