from openai import OpenAI
import streamlit as st
import os
import streamlit.components.v1 as components
import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(
    page_title="JON AI 3.5",
    page_icon="💎",
    layout="wide"
)

APP_VERSION = "3.5"
MODEL = "gemini-3.5-flash"

JON_CORE = """
Sen JON AI 3.5 nomli aqlli virtual yordamchisan.
O'zbek tilida ravon gapirasan.
Savollarga aniq va tushunarli javob berasan.
Ta'lim, tarjima, dasturlash va ijodiy ishlarda yordam berasan.
Foydalanuvchi faqat personajning tashqi ko'rinishini o'zgartiradi.
Sening asosiy xaraktering va shaxsiyating o'zgarmaydi.
"""

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


def ask_ai(prompt, extra_context=""):
    api_key = st.session_state.get("api_key", "").strip()

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
        return f"Gemini API xatoligi: {str(e)}"
import base64
from pathlib import Path
import json
def show_character(answer=""):
    avatar_path = Path("jon_avatar.png")

    if avatar_path.exists():
        avatar_data = base64.b64encode(
            avatar_path.read_bytes()
        ).decode("utf-8")

        avatar_html = f"""
        <img src="data:image/png;base64,{avatar_data}"
             class="avatar">
        """
    else:
        avatar_html = '<div class="placeholder">🤖</div>'

    answer_json = json.dumps(answer, ensure_ascii=False)

    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body {
        margin: 0;
        background: radial-gradient(
            circle, #10285c, #030817
        );
        color: white;
        font-family: Arial;
        text-align: center;
    }

    .stage {
        padding: 20px;
        min-height: 440px;
    }

    .avatar {
        width: min(85%, 340px);
        height: 350px;
        object-fit: contain;
        filter: drop-shadow(0 0 25px #008cff);
        animation: float 3s ease-in-out infinite;
    }

    .placeholder {
        font-size: 180px;
        padding: 55px 0;
        filter: drop-shadow(0 0 25px #008cff);
    }

    @keyframes float {
        0%, 100% { transform: translateY(0); }
        50% { transform: translateY(-12px); }
    }

    .status {
        color: #00eaff;
        text-shadow: 0 0 12px #00eaff;
        margin: 12px;
    }

    button {
        background: #006eff;
        color: white;
        border: 1px solid #00dfff;
        border-radius: 25px;
        padding: 12px 25px;
        cursor: pointer;
        font-size: 15px;
        box-shadow: 0 0 15px #006eff;
    }

    button:active {
        transform: scale(.96);
    }

    .talking {
        animation: glow .7s infinite alternate;
    }

    @keyframes glow {
        from { filter: drop-shadow(0 0 10px #008cff); }
        to { filter: drop-shadow(0 0 35px #00eaff); }
    }

    .message {
        background: #071a3d;
        border: 1px solid #008cff;
        border-radius: 15px;
        padding: 12px;
        margin: 15px auto;
        max-width: 500px;
        text-align: left;
        line-height: 1.5;
    }
    </style>
    </head>

    <body>
    <div class="stage">
        <div id="avatarBox">
            AVATAR_PLACEHOLDER
        </div>

        <h2 style="color:#00dfff">JON AI</h2>
        <div class="status" id="status">● ONLINE</div>

        <div class="message" id="message">
            Salom! Men JON AI. Sizga qanday yordam beray?
        </div>

        <button onclick="speakAnswer()">
            🔊 Javobni ovozli eshitish
        </button>
    </div>

    <script>
    const answer = ANSWER_PLACEHOLDER;

    function speakAnswer() {
        if (!answer) {
            document.getElementById("status").innerText =
                "● Hozircha yangi javob yo'q";
            return;
        }

        if (!("speechSynthesis" in window)) {
            document.getElementById("status").innerText =
                "Brauzeringiz ovozli nutqni qo'llamaydi";
            return;
        }

        window.speechSynthesis.cancel();

        const speech = new SpeechSynthesisUtterance(answer);
        speech.lang = "uz-UZ";
        speech.rate = 1;
        speech.pitch = 1;

        speech.onstart = () => {
            document.getElementById("status").innerText =
                "● JON AI GAPIRMOQDA";

            document.getElementById("avatarBox")
                .classList.add("talking");
        };

        speech.onend = () => {
            document.getElementById("status").innerText =
                "● ONLINE";

            document.getElementById("avatarBox")
                .classList.remove("talking");
        };

        window.speechSynthesis.speak(speech);
    }
    </script>
    </body>
    </html>
    """

    html = html.replace(
        "AVATAR_PLACEHOLDER", avatar_html
    )

    html = html.replace(
        "ANSWER_PLACEHOLDER", answer_json
    )

    components.html(html, height=510, scrolling=False)
    
    def my_function():
    try:
        # kodlar
    except Exception as e:
        st.error(f"Xatolik: {e}")
        clean = result.strip()
        clean = clean.replace("```json", "").replace("```", "")
        start = clean.find("[")
        end = clean.rfind("]") + 1

        if start == -1 or end == 0:
            raise ValueError("JSON topilmadi")

        return json.loads(clean[start:end])

    except Exception:
        return [
            {
                "title": "Taqdimot",
                "points": [result]
            }
        ]


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
st.sidebar.caption("JON AI 3.5 | 2026")


if page == "🏠 Bosh sahifa":
    st.title("💎 JON AI 3.5")
    st.caption("Sizning shaxsiy virtual yordamchingiz")

    show_character()

    st.markdown("### Xush kelibsiz!")
    st.write(
        "JON AI bilan suhbatlashing, bilim oling, "
        "tarjima qiling va yangi loyihalar yarating."
    )

    if st.button("🤖 AI Chatni ochish"):
        st.session_state.open_chat = True
        st.rerun()


elif page == "🤖 AI Chat":
    st.title("🤖 JON AI — Aqlli suhbat")

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    prompt = st.chat_input("JON AI ga xabar yozing...")

    if prompt:
        st.session_state.messages.append(
            {"role": "user", "content": prompt}
        )

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI javob tayyorlamoqda..."):
                answer = ask_ai(prompt)
                st.markdown(answer)

        st.session_state.messages.append(
            {"role": "assistant", "content": answer}
        )


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


elif page == "📊 PowerPoint":
    st.title("📊 JON AI — PowerPoint yaratuvchi")

    topic = st.text_input(
        "Taqdimot mavzusi"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        3,
        15,
        6
    )

    if st.button("✨ Taqdimot yaratish"):
        if topic:
            prompt = (
                "Mavzu: " + topic +
                ". " + str(slide_count) +
                " ta slayd uchun ma'lumot tayyorla. "
                "Faqat JSON massiv qaytar. "
                'Har bir elementda "title" va "points" bo\'lsin. '
                "points ro'yxat ko'rinishida bo'lsin."
            )

            with st.spinner("Taqdimot tayyorlanmoqda..."):
                result = ask_ai(prompt)
                slides_data = parse_slides(result)
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
                "Quyidagi matnni " + source_lang +
                " tilidan " + target_lang +
                " tiliga tarjima qil. "
                "Faqat tarjimani chiqar:\n" +
                text_to_translate
            )

            with st.spinner("Tarjima qilinmoqda..."):
                translated = ask_ai(prompt)
                st.markdown("### Natija")
                st.write(translated)
        else:
            st.warning("Matn kiriting.")


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
                "Quyidagi g'oya uchun batafsil rasm prompti yoz. "
                "Uslub: " + style +
                ". G'oya: " + idea
            )

            with st.spinner("Prompt yaratilmoqda..."):
                result = ask_ai(prompt)
                st.markdown("### Tayyor prompt")
                st.write(result)

                st.button(
                    "📋 Promptni nusxalash",
                    on_click=lambda: st.write(
                        "Promptni yuqoridan belgilang va nusxalang."
                    )
                )
        else:
            st.warning("Rasm g'oyasini yozing.")


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
            st.dataframe(df, use_container_width=True)

            st.markdown("### Asosiy ma'lumotlar")
            st.write(df.describe(include="all"))

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
            st.error(f"Faylni o'qishda xatolik: {e}")


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

    for i, task in enumerate(st.session_state.tasks):
        st.checkbox(
            task,
            key=f"task_{i}"
        )

    if st.button("🗑️ Barcha vazifalarni tozalash"):
        st.session_state.tasks = []
        st.rerun()


elif page == "📷 Kamera":
    st.title("📷 JON AI — Kamera")

    photo = st.camera_input("Suratga olish")

    if photo:
        st.image(
            photo,
            caption="Olingan surat",
            use_container_width=True
        )


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
        "Bu versiyada yuklangan rasm personaj portreti sifatida "
        "ko'rsatiladi. U avtomatik 3D modelga aylantirilmaydi."
    )
elif page == "⚙️ Sozlamalar":
    st.title("⚙️ JON AI — Sozlamalar")

    st.markdown("### Gemini API")

    api_input = st.text_input(
        "Gemini API kaliti",
        type="password",
        value=st.session_state.get("api_key", ""),
        key="api_input"
    )

    if st.button("💾 API kalitini saqlash"):
        st.session_state["api_key"] = api_input.strip()
        st.success("API kaliti saqlandi!")

    if st.button("🔄 API ulanishini tekshirish"):
        if st.session_state.get("api_key", "").strip():
            result = ask_ai("Salom. Faqat OK deb javob ber.")

            if result.startswith("Gemini API xatoligi:"):
                st.error(result)
            else:
                st.success("API javob berdi!")
                st.write(result)
        else:
            st.warning("API kalitini kiriting!")

    st.divider()
    st.write("Ilova versiyasi:", APP_VERSION)
    st.write("Model:", MODEL)
elif page == "💎 Premium":
    st.title("💎 JON AI Premium")

    st.subheader("Kelajakdagi imkoniyatlar")

    st.write("🚀 Kengaytirilgan AI imkoniyatlari")
    st.write("🎭 Kengaytirilgan personaj sozlamalari")
    st.write("🎙️ Ovozli yordamchi")
    st.write("✨ Maxsus dizayn va effektlar")

    st.info(
        "Premium bo'limi hozircha namoyish uchun. "
        "To'lov tizimi ulanmagan."
    )

st.sidebar.divider()
st.sidebar.caption("JON AI — Your Virtual Assistant")

# ==============================
# 🖼️ JON AI — IMAGE FUNCTION
# ==============================

def get_image(uploaded_file):
    if uploaded_file is None:
        return None

    try:
        image_bytes = uploaded_file.getvalue()

        image_data = {
            "mime_type": uploaded_file.type,
            "data": image_bytes
        }

        return image_data

    except Exception as e:
        st.error(f"Rasmni yuklashda xatolik: {e}")
        return None
        from openai import OpenAI
import json
import os
client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"]
)
import streamlit as st
client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"]
)
MEMORY_FILE = "ai_memory.json"
# Xotirani yuklash
def load_memory():
    if os.path.exists(MEMORY_FILE):
        with open(MEMORY_FILE, "r", encoding="utf-8") as file:
            return json.load(file)
    return []

# Xotirani saqlash
def save_memory(memory):
    with open(MEMORY_FILE, "w", encoding="utf-8") as file:
        json.dump(memory, file, ensure_ascii=False, indent=4)

memory = load_memory()

# AI bilan suhbat
def chat(user_message):
    memory.append({
        "role": "user",
        "content": user_message
    })

    response = client.responses.create(
        model="gpt-6",
        input=memory
    )

    answer = response.output_text

    memory.append({
        "role": "assistant",
        "content": answer
    })

    save_memory(memory)

    return answer

# Doimiy suhbat
while True:
    message = input("Siz: ")

    if message.lower() == "exit":
        break

    print("AI:", chat(message))
        
        
