import random
import json
from io import BytesIO
import streamlit as st
import pandas as pd
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="JON AI", page_icon="🤖", layout="wide")
ANIMS = [("portal", "ENERGY PORTAL"), ("particles", "DIGITAL PARTICLES"), ("scan", "LASER SCAN"), ("galaxy", "GALAXY AWAKENING"), ("cube", "HOLOGRAM FORMATION"), ("electric", "ELECTRIC ENERGY"), ("awakening", "JON AI AWAKENING")]
if "intro" not in st.session_state: st.session_state.intro = random.choice(ANIMS)
if "messages" not in st.session_state: st.session_state.messages = []
if "tasks" not in st.session_state: st.session_state.tasks = []
if "api_key" not in st.session_state: st.session_state.api_key = ""


def build_powerpoint(topic, slides_data, theme_name):
    """Create a downloadable PowerPoint presentation."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    themes = {
        "Moviy futuristik": ((5, 16, 38), (0, 217, 255), (235, 247, 255)),
        "Binafsha zamonaviy": ((25, 13, 48), (190, 130, 255), (248, 240, 255)),
        "Oq minimal": ((248, 250, 252), (20, 90, 150), (25, 35, 50)),
        "Yashil akademik": ((10, 40, 35), (80, 220, 170), (235, 255, 245)),
    }
    bg, accent, foreground = themes[theme_name]

    def set_background(slide):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg)

    def add_text(slide, text, left, top, width, height, size, color, bold=False):
        box = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        frame = box.text_frame
        frame.word_wrap = True
        frame.margin_left = Inches(0.08)
        frame.margin_right = Inches(0.08)
        frame.margin_top = Inches(0.04)
        p = frame.paragraphs[0]
        p.text = str(text)
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = RGBColor(*color)
        return box

    # Cover slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_text(slide, "JON AI PRESENTATION", 0.8, 1.2, 11.7, 0.6, 18, accent, True)
    add_text(slide, topic, 0.8, 2.3, 11.7, 2.0, 34, foreground, True)
    add_text(slide, "Sun’iy intellekt yordamida tayyorlandi", 0.9, 5.9, 11.5, 0.5, 15, accent)

    # Content slides
    for number, item in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_background(slide)
        title = item.get("title", f"{number}-slayd")
        bullets = item.get("bullets", [])
        add_text(slide, f"{number:02d}  /  {title}", 0.65, 0.55, 12.0, 0.85, 27, accent, True)

        if isinstance(bullets, str):
            bullets = [line.strip("- •") for line in bullets.splitlines() if line.strip()]
        bullets = bullets[:6] if bullets else ["Ushbu bo‘lim uchun ma’lumot qo‘shing."]
        y = 1.8
        for bullet in bullets:
            add_text(slide, "•  " + str(bullet), 1.0, y, 11.3, 0.65, 20, foreground)
            y += 0.82

        add_text(slide, "JON AI", 11.2, 7.0, 1.4, 0.25, 10, accent, True)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def parse_presentation_json(raw_text, expected_count):
    """Parse model JSON; return safe fallback slides if formatting differs."""
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.split("\\n", 1)[-1]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()
    try:
        data = json.loads(cleaned)
        slides = data.get("slides", data) if isinstance(data, dict) else data
        if isinstance(slides, list) and slides:
            normalized = []
            for item in slides[:expected_count]:
                if isinstance(item, dict):
                    bullets = item.get("bullets", item.get("points", []))
                    if isinstance(bullets, str):
                        bullets = [x.strip("- •") for x in bullets.splitlines() if x.strip()]
                    normalized.append({
                        "title": str(item.get("title", "Slayd")),
                        "bullets": [str(x) for x in bullets[:6]]
                    })
            if normalized:
                return normalized
    except Exception:
        pass

    # If the model did not return JSON, retain its text in readable slide sections.
    chunks = [part.strip() for part in raw_text.split("\\n\\n") if part.strip()]
    fallback = []
    for i, chunk in enumerate(chunks[:expected_count]):
        lines = [line.strip(" -•\\t") for line in chunk.splitlines() if line.strip()]
        fallback.append({
            "title": lines[0][:100] if lines else f"{i+1}-slayd",
            "bullets": lines[1:6] or [chunk[:350]]
        })
    return fallback or [{"title": "Asosiy ma’lumot", "bullets": [raw_text[:900]]}]

st.markdown(r"""
<style>
.stApp{background:radial-gradient(circle at 50% 0%,#153b70 0%,transparent 38%),linear-gradient(135deg,#050816,#0b1734 55%,#17113b);color:#f4f8ff}
.block-container{padding-top:1.5rem;max-width:1200px}h1,h2,h3{color:#8ceeff}
.hero{text-align:center;font-size:clamp(36px,6vw,64px);font-weight:900;letter-spacing:4px;color:#8ceeff;text-shadow:0 0 12px #00cfff,0 0 32px #006dff}
.sub{text-align:center;color:#c7d8f5;font-size:17px}
.holo{height:300px;display:flex;align-items:center;justify-content:center;flex-direction:column;border:1px solid #00d9ff;border-radius:24px;background:radial-gradient(circle,#123b70,#071329 72%);box-shadow:0 0 30px #007bff55;margin:18px 0;overflow:hidden}
.head{width:76px;height:82px;border:3px solid #00eaff;border-radius:45% 45% 40% 40%;position:relative;background:#00d9ff22;box-shadow:0 0 25px #00d9ff;animation:float 3s infinite}
.head:after{content:'◉　◉';position:absolute;top:24px;left:10px;color:#baffff;font-size:17px;text-shadow:0 0 12px cyan}
.body{width:145px;height:80px;margin-top:8px;border:3px solid #00d9ff;border-radius:65px 65px 15px 15px;background:#00eaff33;box-shadow:0 0 25px #00d9ff;animation:float 3s infinite}
.ring{width:190px;height:34px;border:2px solid #00eaff;border-radius:50%;box-shadow:0 0 18px #00d9ff;margin-top:-8px}.name{color:#8ffaff;font-size:19px;font-weight:bold;margin-top:12px;text-shadow:0 0 12px #00d9ff}
@keyframes float{50%{transform:translateY(-10px)}}
@keyframes portal{0%{transform:scale(.1) rotate(-180deg);opacity:0}70%{transform:scale(1.08)}100%{transform:scale(1);opacity:1}}
@keyframes particles{0%{transform:translateY(35px) scale(.5);opacity:0}100%{transform:none;opacity:1}}
@keyframes scan{0%{filter:brightness(.2);clip-path:inset(0 0 100% 0)}100%{filter:brightness(1);clip-path:inset(0)}}
@keyframes galaxy{0%{transform:rotate(-12deg) scale(.65);opacity:0}100%{transform:none;opacity:1}}
@keyframes cube{0%{transform:rotateY(90deg) scale(.4);opacity:0}100%{transform:none;opacity:1}}
@keyframes electric{0%,100%{filter:brightness(.7)}25%{filter:brightness(1.8)}50%{filter:brightness(.9)}75%{filter:brightness(1.6)}}
@keyframes awakening{0%{transform:translateY(45px);opacity:0}100%{transform:none;opacity:1}}
.anim-portal{animation:portal 1.4s both}.anim-particles{animation:particles 1.4s both}.anim-scan{animation:scan 1.5s both}.anim-galaxy{animation:galaxy 1.5s both}.anim-cube{animation:cube 1.5s both}.anim-electric{animation:electric 1.5s both}.anim-awakening{animation:awakening 1.5s both}
.feature{background:#071d3d99;border:1px solid #176a9c;border-radius:16px;padding:14px;min-height:100px}
.stButton>button{background:linear-gradient(90deg,#007bff,#702cff);color:white;border:1px solid #00d9ff;border-radius:12px;min-height:42px}
.stChatMessage{background:#00234b59;border:1px solid #00bde8;border-radius:15px}
</style>
""", unsafe_allow_html=True)

def ai(prompt, system="Sen JON AI nomli yordamchisan. O'zbek tilida aniq, samimiy va foydali javob ber."):
    key = st.session_state.api_key.strip()
    if not key:
        st.warning("Avval chap tomondagi sozlamalarga Gemini API kalitini kiriting.")
        return None
    try:
        client = genai.Client(api_key=key)
        response = client.models.generate_content(model="gemini-3.5-flash-lite", contents=[system, prompt])
        return response.text or "Hozircha javob olinmadi."
    except Exception as e:
        st.error("AI bilan bog'lanishda xatolik: " + str(e))
        return None

def hologram():
    cls, label = st.session_state.intro
    st.markdown(f'''<div class="holo"><div class="anim-{cls}"><div class="head"></div><div class="body"></div><div class="ring"></div><div class="name">JON AI</div></div></div><p class="sub">✦ {label} · SYSTEM ACTIVATED ✦</p>''', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 💎 JON AI")
    st.caption("Sizning shaxsiy AI yordamchingiz")
    st.session_state.api_key = st.text_input("Gemini API kaliti", type="password", value=st.session_state.api_key, help="Kalitni hech kimga yubormang.")
    st.divider()
    page = st.radio("BO'LIMLAR", ["🏠 Bosh sahifa", "🤖 AI suhbat", "🎓 Ta'lim markazi", "📝 Prezentatsiya", "🌐 Tarjimon", "🎨 Surat g'oyasi", "🎵 Kayfiyat va musiqa", "📊 Excel tahlil", "📅 Kunlik rejalar", "📷 Kamera", "💎 Premium"])
    if st.button("🗑️ Suhbatni tozalash", use_container_width=True):
        st.session_state.messages = []; st.rerun()

if page == "🏠 Bosh sahifa":
    st.markdown('<div class="hero">JON AI</div><p class="sub">Kelajak yordamchingiz — barcha imkoniyatlar bir joyda</p>', unsafe_allow_html=True)
    hologram()
    st.subheader("✨ Imkoniyatlar markazi")
    cards = [("🤖","AI suhbat","Savollar va kundalik yordam"),("🎓","Ta'lim markazi","IELTS va Multilevel"),("📝","Prezentatsiya","Slayd rejasi"),("🌐","Tarjimon","Ko'p tilli tarjima"),("🎨","Surat laboratoriyasi","Rasm uchun prompt"),("🎵","Musiqa","Kayfiyatga mos tavsiya"),("📊","Excel","Jadval tahlili"),("📅","Rejalar","Kunlik vazifalar"),("📷","Kamera","Suratga olish"),("💎","Premium","Kelajakdagi imkoniyatlar")]
    cols=st.columns(2)
    for i,(emoji,title,desc) in enumerate(cards):
        with cols[i%2]: st.markdown(f'<div class="feature"><h3>{emoji} {title}</h3>{desc}</div>',unsafe_allow_html=True)
    st.info("Kerakli bo'limni chap menyudan tanlang.")

elif page == "🤖 AI suhbat":
    st.title("🤖 JON AI bilan suhbat")
    for m in st.session_state.messages:
        with st.chat_message(m["role"]): st.markdown(m["content"])
    prompt=st.chat_input("JON AI ga savolingizni yozing...")
    if prompt:
        st.session_state.messages.append({"role":"user","content":prompt})
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("JON AI o'ylamoqda..."):
                answer=ai(prompt)
                if answer:
                    st.markdown(answer); st.session_state.messages.append({"role":"assistant","content":answer})

elif page == "🎓 Ta'lim markazi":
    st.title("🎓 IELTS va Multilevel")
    course=st.selectbox("Dars turi",["IELTS Speaking","IELTS Writing","IELTS Reading","IELTS Listening","Multilevel grammatika","Multilevel lug'at"])
    level=st.selectbox("Daraja",["Beginner","Elementary","Intermediate","Upper-Intermediate","Advanced"])
    topic=st.text_input("Mavzu", "Speaking Part 2")
    if st.button("📚 Darsni boshlash"):
        with st.spinner("Dars tayyorlanmoqda..."):
            result=ai(f"{course} bo'yicha {level} darajasiga mos {topic} mavzusida dars, misollar, mashqlar va javoblar tayyorla.","Sen tajribali ingliz tili o'qituvchisisan. Sodda va bosqichma-bosqich tushuntir.")
        if result: st.markdown(result)

elif page == "📝 Prezentatsiya":
    st.title("📝 JON AI — PowerPoint Maker")
    st.write("Mavzuni kiriting: JON AI slaydlar matnini tuzadi va yuklab olinadigan PowerPoint faylini yaratadi.")

    topic = st.text_input("Prezentatsiya mavzusi", placeholder="Masalan: Sun’iy intellektning ta’limdagi o‘rni")
    col1, col2 = st.columns(2)
    with col1:
        count = st.slider("Slaydlar soni", 3, 20, 8)
        lang = st.selectbox("Taqdimot tili", ["O'zbek", "English", "Русский"])
    with col2:
        theme = st.selectbox(
            "Dizayn uslubi",
            ["Moviy futuristik", "Binafsha zamonaviy", "Oq minimal", "Yashil akademik"]
        )
        audience = st.selectbox("Kimlar uchun?", ["Maktab o'quvchilari", "Talabalar", "O'qituvchilar", "Umumiy auditoriya"])

    if st.button("✨ PowerPoint yaratish", use_container_width=True):
        if not topic.strip():
            st.warning("Avval prezentatsiya mavzusini kiriting.")
        elif not st.session_state.get("api_key", "").strip():
            st.warning("Avval chap tomondagi sozlamalarga Gemini API kalitini kiriting.")
        else:
            prompt = f"""
            '{topic}' mavzusida {count} ta mazmunli slayd tayyorla.
            Til: {lang}. Auditoriya: {audience}.
            Faqat quyidagi JSON formatida javob ber:
            {{
              "slides": [
                {{"title": "Slayd sarlavhasi", "bullets": ["qisqa fikr 1", "qisqa fikr 2", "qisqa fikr 3"]}}
              ]
            }}
            Aynan {count} ta slayd bo'lsin. Har bir slaydda 3-5 ta qisqa, aniq punkt bo'lsin.
            Birinchi slayd mavzuni tanishtirsin, oxirgi slayd xulosa va asosiy fikrlarni bersin.
            JSONdan tashqari izoh yozma.
            """
            with st.spinner("JON AI slaydlarni tayyorlamoqda..."):
                raw = ai(prompt, "Sen professional taqdimot muallifisan. Tuzilmali, aniq va tekshiriladigan ma'lumot ber.")
            if raw:
                slide_items = parse_presentation_json(raw, count)
                pptx_bytes = build_powerpoint(topic, slide_items, theme)
                st.success("PowerPoint taqdimotingiz tayyor!")
                st.subheader("Slaydlar ko'rinishi")
                for idx, slide_item in enumerate(slide_items, start=1):
                    with st.expander(f"{idx}. {slide_item['title']}", expanded=(idx == 1)):
                        for bullet in slide_item["bullets"]:
                            st.write("• " + bullet)
                safe_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in topic).strip("_")
                st.download_button(
                    "⬇️ PowerPoint (.pptx) yuklab olish",
                    data=pptx_bytes,
                    file_name=(safe_name[:50] or "jon_ai_presentation") + ".pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

elif page == "🌐 Tarjimon":
    st.title("🌐 JON AI — Universal tarjimon")
    st.write("Til nomini o'zingiz yozing — faqat ro'yxat bilan cheklanmaydi.")

    source_lang = st.selectbox(
        "Matn qaysi tilda?",
        ["Avtomatik aniqlash", "O'zbek", "English", "Русский", "العربية",
         "中文", "日本語", "한국어", "Français", "Deutsch", "Español", "Boshqa"]
    )
    target_choice = st.selectbox(
        "Tez tanlash — tarjima qilinadigan til",
        ["Tilni o'zim yozaman", "O'zbek", "English", "Русский", "العربية",
         "中文 (Xitoy)", "日本語 (Yapon)", "한국어 (Koreys)", "Français (Fransuz)",
         "Deutsch (Nemis)", "Español (Ispan)", "Português (Portugal/Braziliya)",
         "Türkçe (Turk)", "Italiano (Italyan)", "हिन्दी (Hind)", "اردو (Urdu)",
         "فارسی (Fors)", "Қазақша (Qozoq)", "Кыргызча (Qirg'iz)", "Тоҷикӣ (Tojik)",
         "Українська (Ukrain)", "Polski (Polyak)", "Nederlands (Golland)",
         "עברית (Ibroniy)", "ไทย (Tay)", "Tiếng Việt (Vyetnam)", "বাংলা (Bengal)",
         "Latin (Lotin)", "Ancient Greek (Qadimgi yunon)", "Sanskrit"]
    )
    target_custom = ""
    if target_choice == "Tilni o'zim yozaman":
        target_custom = st.text_input(
            "Til nomini kiriting",
            placeholder="Masalan: Suahili, Amhar, Maori yoki boshqa til"
        )
    target_lang = target_custom.strip() if target_choice == "Tilni o'zim yozaman" else target_choice

    style = st.selectbox(
        "Tarjima uslubi",
        ["Tabiiy va ravon", "So'zma-so'z", "Rasmiy", "Akademik", "Oddiy va tushunarli"]
    )
    explain = st.checkbox("Tarjima bilan birga muhim so'zlarni va iboralarni izohla", value=False)
    text_to_translate = st.text_area("Tarjima qilinadigan matn", height=180,
                                     placeholder="Matnni shu yerga kiriting...")
    if st.button("🌍 Tarjima qilish", use_container_width=True):
        if not text_to_translate.strip():
            st.warning("Avval tarjima qilinadigan matnni kiriting.")
        elif not target_lang:
            st.warning("Tarjima tilini yozing yoki ro'yxatdan tanlang.")
        else:
            source_instruction = (
                "Manba tilini o'zing aniqlagin."
                if source_lang == "Avtomatik aniqlash"
                else f"Manba tili: {source_lang}."
            )
            extra = (
                "Tarjimadan keyin asosiy so'z va iboralarni qisqacha izohla."
                if explain else
                "Faqat tarjimani ber; zarur bo'lsa, noaniq joyga qisqa eslatma qo'sh."
            )
            prompt = f"""
{source_instruction}
Matnni {target_lang} tiliga {style.lower()} uslubda tarjima qil.
{extra}
Mazmun, ohang, ism-shariflar va formatni imkon qadar saqla.
Agar bu til yoki yozuv bo'yicha ishonchli tarjima qilish qiyin bo'lsa,
buni ochiq ayt; taxminiy iboralarni aniq tarjima deb ko'rsatma.
Qadimgi yoki kam resursli tillarda noaniqliklarni alohida belgilagin.
Matn:
{text_to_translate}
"""
            with st.spinner("Tarjima tayyorlanmoqda..."):
                result = ai(prompt, "Sen ko'p tilli tarjimon yordamchisan. Aniqlikni birinchi o'ringa qo'y.")
            if result:
                st.markdown("### Natija")
                st.markdown(result)
                st.caption("Eslatma: juda kam ishlatiladigan, qadimgi yoki yo'qolib ketgan tillarda tarjima xatolari bo'lishi mumkin. Muhim hujjatlarni mutaxassisga tekshirtiring.")

elif page == "🎨 Surat g'oyasi":
    st.title("🎨 Surat laboratoriyasi")
    desc=st.text_area("Qanday surat kerak?")
    style=st.selectbox("Uslub",["Realistik","3D animatsiya","Anime","Cyberpunk","Minimalistik"])
    if st.button("🪄 Rasm promptini yaratish"):
        result=ai(f"{desc} uchun {style} uslubida batafsil inglizcha image-generation prompt yoz va o'zbekcha izoh ber.")
        if result: st.markdown(result)
    st.info("Bu versiya rasm yaratish promptini tayyorlaydi. Haqiqiy rasm generatsiyasi uchun alohida rasm API ulash kerak.")

elif page == "🎵 Kayfiyat va musiqa":
    st.title("🎵 Kayfiyatga mos musiqa")
    mood=st.selectbox("Kayfiyatingizni tanlang",["Quvnoq","Tinchlanmoqchiman","Charchaganman","Diqqatni jamlash","Ilhom kerak","Aniq bilmayman"])
    if st.button("🎧 Tavsiya olish"):
        result=ai(f"Kayfiyatim: {mood}. Mos musiqa janrlari, qonuniy platformalarda qidirish so'zlari va 3 ta foydali faoliyat tavsiya qil.")
        if result: st.markdown(result)
    st.caption("Musiqa avtomatik ijro etilmaydi; hozircha tavsiyalar beriladi.")

elif page == "📊 Excel tahlil":
    st.title("📊 Excel va hisob-kitob")
    file=st.file_uploader("Excel yoki CSV faylini yuklang",type=["xlsx","xls","csv"])
    if file:
        try:
            df=pd.read_csv(file) if file.name.lower().endswith(".csv") else pd.read_excel(file)
            st.success(f"Yuklandi: {df.shape[0]} qator, {df.shape[1]} ustun")
            st.dataframe(df,use_container_width=True)
            st.subheader("Sonli ustunlar statistikasi")
            st.dataframe(df.describe(),use_container_width=True)
            q=st.text_input("Jadval bo'yicha savol")
            if st.button("AI bilan tahlil"):
                result=ai(f"Jadval namunasi:\n{df.head(30).to_string()}\nSavol: {q or 'Asosiy xulosalarni chiqar'}")
                if resu
    
