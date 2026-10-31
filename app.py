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
    page_title="JON AI",
    page_icon="💎",
    layout="wide"
)

MODEL = "gemini-2.5-flash"

if "messages" not in st.session_state:
    st.session_state.messages = []

if "tasks" not in st.session_state:
    st.session_state.tasks = []

if "character_name" not in st.session_state:
    st.session_state.character_name = "JON"

if "character_style" not in st.session_state:
    st.session_state.character_style = "Hologram"

if "avatar" not in st.session_state:
    st.session_state.avatar = None

if "api_key" not in st.session_state:
    st.session_state.api_key = ""

st.markdown("""
<style>
.stApp {
    background: radial-gradient(
        circle at 50% 0%,
        #173d79 0%,
        #08152e 45%,
        #030711 100%
    );
    color: white;
}

h1, h2, h3, p, label {
    color: white !important;
}

section[data-testid="stSidebar"] {
    background: #071329;
    border-right: 1px solid #00d9ff;
}

.stButton > button {
    background: linear-gradient(90deg,#006eff,#762cff);
    color: white;
    border: 1px solid #00d9ff;
    border-radius: 13px;
    padding: 10px 20px;
    font-weight: bold;
    width: 100%;
}

.stButton > button:hover {
    border: 1px solid white;
    box-shadow: 0 0 18px #00d9ff;
}

.holo {
    text-align: center;
    padding: 25px;
    border: 1px solid #00d9ff;
    border-radius: 25px;
    background: rgba(0, 100, 255, 0.09);
    box-shadow: 0 0 35px rgba(0,217,255,.25);
}

.avatar-head {
    width: 95px;
    height: 95px;
    margin: auto;
    border-radius: 50%;
    background: radial-gradient(circle,#b9ffff,#00aaff,#173dff);
    box-shadow: 0 0 35px #00d9ff;
    animation: pulse 2s infinite alternate;
}

.avatar-body {
    width: 145px;
    height: 120px;
    margin: 15px auto;
    border-radius: 60% 60% 20% 20%;
    background: linear-gradient(140deg,#00d9ff,#1646bb,#772cff);
    box-shadow: 0 0 30px #007bff;
}

.ring {
    width: 180px;
    height: 35px;
    border: 3px solid #00d9ff;
    border-radius: 50%;
    margin: 15px auto;
    box-shadow: 0 0 20px #00d9ff;
}

@keyframes pulse {
    from {transform: scale(.95); opacity:.7;}
    to {transform: scale(1.08); opacity:1;}
}

.card {
    padding: 18px;
    border: 1px solid #168dff;
    border-radius: 16px;
    background: rgba(0,80,180,.15);
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)


def ask_ai(prompt):
    key = st.session_state.api_key

    if not key:
        return "Sozlamalar bo‘limiga Gemini API kalitini kiriting."

    try:
        client = genai.Client(api_key=key)

        response = client.models.generate_content(
            model=MODEL,
            contents=prompt
        )

        return response.text or "Javob olinmadi."

    except Exception as e:
        return f"AI xatosi: {e}"


def show_hologram():
    st.markdown('<div class="holo">', unsafe_allow_html=True)

    if st.session_state.avatar:
        st.image(st.session_state.avatar, width=180)
    else:
        st.markdown("""
        <div class="avatar-head"></div>
        <div class="avatar-body"></div>
        <div class="ring"></div>
        """, unsafe_allow_html=True)

    st.markdown(
        f"<h2>{st.session_state.character_name}</h2>",
        unsafe_allow_html=True
    )

    st.markdown(
        f"<p>✦ {st.session_state.character_style} SYSTEM ACTIVE ✦</p>",
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)# =====================================
# 5. INTERNETDAN RASM QIDIRISH
# =====================================

def get_image(query):
    try:
        url = "https://commons.wikimedia.org/w/api.php"

        params = {
            "action": "query",
            "generator": "search",
            "gsrsearch": query + " filetype:bitmap",
            "gsrnamespace": 6,
            "gsrlimit": 5,
            "prop": "imageinfo",
            "iiprop": "url",
            "iiurlwidth": 900,
            "format": "json"
        }

        headers = {
            "User-Agent": "JONAI-Presentation/1.0"
        }

        response = requests.get(
            url,
            params=params,
            headers=headers,
            timeout=15
        )

        data = response.json()
        pages = data.get("query", {}).get("pages", {})

        for page in pages.values():
            info = page.get("imageinfo", [{}])[0]
            image_url = info.get("thumburl") or info.get("url")

            if image_url:
                image_response = requests.get(
                    image_url,
                    headers=headers,
                    timeout=15
                )

                if image_response.status_code == 200:
                    return image_response.content

    except Exception:
        return None

    return None


# =====================================
# 6. POWERPOINT YARATISH
# =====================================

def create_presentation(topic, slide_data, theme, add_images):

    prs = Presentation()

    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    colors = {
        "Ko‘k": (10, 35, 90),
        "Qora": (15, 15, 25),
        "Oq": (245, 245, 245),
        "Binafsha": (55, 20, 90)
    }

    bg = colors.get(theme, colors["Ko‘k"])

    for index, item in enumerate(slide_data):

        slide = prs.slides.add_slide(
            prs.slide_layouts[6]
        )

        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg)

        title = item.get(
            "title",
            f"Slayd {index + 1}"
        )

        points = item.get("points", [])

        title_box = slide.shapes.add_textbox(
            Inches(.6),
            Inches(.4),
            Inches(12),
            Inches(.8)
        )

        title_frame = title_box.text_frame
        title_frame.text = title

        for paragraph in title_frame.paragraphs:
            paragraph.font.size = Pt(28)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor(
                0, 217, 255
            )

        text_box = slide.shapes.add_textbox(
            Inches(.8),
            Inches(1.6),
            Inches(6),
            Inches(5)
        )

        frame = text_box.text_frame
        frame.word_wrap = True

        for i, point in enumerate(points):

            paragraph = (
                frame.paragraphs[0]
                if i == 0
                else frame.add_paragraph()
            )

            paragraph.text = "• " + str(point)
            paragraph.font.size = Pt(19)
            paragraph.font.color.rgb = RGBColor(
                255, 255, 255
            )
            paragraph.space_after = Pt(14)

        # Slaydga mos rasm qo‘shish

        if add_images and index > 0:

            image_bytes = get_image(
                title + " " + topic
            )

            if image_bytes:
                try:
                    image_stream = BytesIO(image_bytes)

                    slide.shapes.add_picture(
                        image_stream,
                        Inches(7.2),
                        Inches(1.7),
                        width=Inches(5.4),
                        height=Inches(4.5)
                    )

                except Exception:
                    pass

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output


# =====================================
# 7. AI JAVOBINI JSON GA AJRATISH
# =====================================

def parse_slides(result):

    clean = result.replace(
        "```json", ""
    ).replace("```", "").strip()

    data = json.loads(clean)

    return data["slides"]


# =====================================
# 8. SIDEBAR MENYUSI
# =====================================

st.sidebar.markdown("# 💎 JON AI")
st.sidebar.caption(
    "YOUR PERSONAL AI ASSISTANT"
)

pages = [
    "🏠 Bosh sahifa",
    "💬 AI suhbat",
    "🎓 Ta'lim markazi",
    "📊 Prezentatsiya",
    "🌐 Tarjimon",
    "🎨 Surat g‘oyasi",
    "📈 Excel tahlil",
    "📅 Kunlik rejalar",
    "📷 Kamera",
    "🤖 Personaj sozlamalari",
    "⚙️ Sozlamalar",
    "💎 Premium"
]

page = st.sidebar.radio(
    "BO‘LIMLAR",
    pages
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    "✦ JON AI SYSTEM"
)


# =====================================
# 9. BOSH SAHIFA
# =====================================

if page == "🏠 Bosh sahifa":

    st.title("💎 JON AI")

    st.caption(
        "Sizning shaxsiy virtual yordamchingiz"
    )

    show_hologram()

    st.markdown("### ✦ Xush kelibsiz!")

    st.write(
        "JON AI — bilim olish, tarjima, tahlil "
        "va ijodiy ishlarda yordam beruvchi "
        "sun’iy intellekt yordamchisi."
    )

    st.info(
        "Yon menyudan kerakli bo‘limni tanlang."
            )

            if explain:
                prompt += """
                Tarjimadan keyin muhim so‘zlar,
                iboralar va grammatik tuzilishlarni
                qisqacha tushuntir.
                """

            with st.spinner("Tarjima qilinmoqda..."):
                translation = ask_ai(prompt)

            st.markdown("### Natija")
            st.write(translation)


# =====================================
# 14. SURAT G‘OYASI
# =====================================

elif page == "🎨 Surat g‘oyasi":

    st.title("🎨 JON AI — Surat g‘oyasi")

    idea = st.text_area(
        "Qanday surat yaratmoqchisiz?"
    )

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Realistik",
            "3D",
            "Anime",
            "Cyberpunk",
            "Hologram",
            "Cartoon",
            "Cinematic"
        ]
    )

    if st.button("Prompt yaratish"):

        if idea.strip():

            prompt = f"""
            Create a professional image generation prompt.

            Idea: {idea}
            Style: {style}

            Include lighting, camera angle, colors,
            environment and visual details.

            Return the prompt in English.
            """

            with st.spinner("Prompt tayyorlanmoqda..."):
                st.write(ask_ai(prompt))

    st.info(
        "Bu bo‘lim rasm yaratish uchun prompt tayyorlaydi. "
        "Haqiqiy rasm yaratish uchun alohida image API kerak."
    )


# =====================================
# 15. EXCEL TAHLIL
# =====================================

elif page == "📈 Excel tahlil":

    st.title("📈 JON AI — Excel tahlil")

    uploaded = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded:

        try:

            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded)
            else:
                df = pd.read_excel(uploaded)

            st.success("Fayl yuklandi!")

            st.dataframe(
                df.head(30),
                use_container_width=True
            )

            st.write("Qatorlar:", len(df))
            st.write("Ustunlar:", len(df.columns))

            if st.button("AI bilan tahlil qilish"):

                sample = df.head(40).to_string()

                prompt = f"""
                Quyidagi jadvalni tahlil qil.
                Asosiy tendensiyalar, xatolar va
                xulosalarni tushunarli izohla.

                {sample}
                """

                with st.spinner("Tahlil qilinmoqda..."):
                    st.write(ask_ai(prompt))

        except Exception as e:
            st.error(f"Faylni o‘qishda xato: {e}")


# =====================================
# 16. KUNLIK REJALAR
# =====================================

elif page == "📅 Kunlik rejalar":

    st.title("📅 JON AI — Kunlik rejalar")

    task = st.text_input("Yangi vazifa")

    if st.button("Vazifa qo‘shish"):

        if task.strip():

            st.session_state.tasks.append({
                "task": task,
                "done": False
            })

            st.rerun()

    st.markdown("### Mening rejalarim")

    for i, item in enumerate(st.session_state.tasks):

        checked = st.checkbox(
            item["task"],
            value=item["done"],
            key=f"task_{i}"
        )

        st.session_state.tasks[i]["done"] = checked

    if st.button("Bajarilganlarni tozalash"):

        st.session_state.tasks = [
            t for t in st.session_state.tasks
            if not t["done"]
        ]

        st.rerun()


# =====================================
# 17. KAMERA
# =====================================

elif page == "📷 Kamera":

    st.title("📷 JON AI — Kamera")

    picture = st.camera_input("Suratga oling")

    if picture:

        st.image(
            picture,
            caption="Siz olgan surat"
        )

        st.info(
            "Rasmni AI orqali tahlil qilish "
            "funksiyasi keyingi bosqichda ulanadi."
        )


# =====================================
# 18. PERSONAJ SOZLAMALARI
# =====================================

elif page == "🤖 Personaj sozlamalari":

    st.title("🤖 JON AI — Personaj")

    name = st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    style = st.selectbox(
        "Personaj ko‘rinishi",
        [
            "Hologram",
            "Cyber",
            "Robot",
            "Minimal"
        ]
    )

    avatar_file = st.file_uploader(
        "Personaj suratini yuklang",
        type=["png", "jpg", "jpeg"]
    )

    if st.button("O‘zgarishlarni saqlash"):

        st.session_state.character_name = name
        st.session_state.character_style = style

        if avatar_file:
            st.session_state.avatar = avatar_file.getvalue()

        st.success("Sozlamalar saqlandi!")
        st.rerun()

    show_hologram()


# =====================================
# 19. API SOZLAMALARI
# =====================================

elif page == "⚙️ Sozlamalar":

    st.title("⚙️ JON AI — Sozlamalar")

    st.markdown("### Gemini API")

    api_key = st.text_input(
        "Gemini API kaliti",
        type="password",
        value=st.session_state.api_key
    )

    if st.button("API kalitini saqlash"):

        st.session_state.api_key = api_key

        st.success("Kalit joriy sessiya uchun saqlandi.")


# =====================================
# 20. PREMIUM
# =====================================

elif page == "💎 Premium":

    st.title("💎 JON AI PREMIUM")

    st.markdown("""
    ### Premium imkoniyatlar

    - Kengaytirilgan AI funksiyalari
    - Ko‘proq prezentatsiya imkoniyatlari
    - Kengaytirilgan tarjima
    - Personaj sozlamalari
    - Kelajakdagi ovozli yordamchi

    """)
    st.info("Premium obuna tizimi hali ulanmagan.")


# =====================================
# 21. YAKUN
# =====================================

st.sidebar.markdown("---")
st.sidebar.caption("JON AI © 2026")
    
