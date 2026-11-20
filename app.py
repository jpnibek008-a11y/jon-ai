# ============================================================
# JON AI LIVE 8.0
# Streamlit + Gemini 3.8 + Three.js
# Professional unified application
# ============================================================

import json
import re
import time
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from google import genai
from pptx import Presentation


# ============================================================
# 1. APP CONFIGURATION
# ============================================================

APP_VERSION = "8.0"
DEFAULT_MODEL = "gemini-3.5-flash"

st.set_page_config(
    page_title="JON AI LIVE",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

PAGES = [
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

JON_CORE = """
Sen JON AI nomli professional virtual yordamchisan.

Asosiy til: o'zbek tili.

Qoidalar:
1. Aniq, ravon va foydali javob ber.
2. Ma'lumotlarni to'qib chiqarma.
3. Ishonching komil bo'lmasa, buni ochiq ayt.
4. Murakkab vazifalarni bosqichma-bosqich tushuntir.
5. Foydalanuvchiga hurmat bilan munosabatda bo'l.
6. Maxfiy kalitlar va parollarni so'рама yoki oshkor qilma.
7. Keraksiz uzun javoblardan qoch.
8. Ta'lim, dasturlash, tarjima, tahlil va ijodiy
   vazifalarda professional yordam ko'rsat.
"""


# ============================================================
# 2. SESSION STATE
# ============================================================

DEFAULTS = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "api_key": "",
    "page": PAGES[0],
    "model": DEFAULT_MODEL,
    "intro_enabled": True
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# 3. DESIGN SYSTEM
# ============================================================

st.markdown("""
<style>

@import url(
'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap'
);

.stApp {
    background:
    radial-gradient(
        ellipse at 10% 0%,
        rgba(0,180,255,.16),
        transparent 42%
    ),
    radial-gradient(
        ellipse at 90% 20%,
        rgba(120,50,255,.13),
        transparent 45%
    ),
    linear-gradient(145deg,#050b16,#10182d);

    color:#edf5ff;
}

html, body, [class*="css"] {
    font-family:Inter,sans-serif;
}

[data-testid="stSidebar"] {
    background:#091426;
    border-right:1px solid rgba(80,180,255,.15);
}

.block-container {
    max-width:1450px;
    padding-top:1.5rem;
    padding-bottom:3rem;
}

h1,h2,h3 {
    color:#eaf7ff;
}

.jon-hero {
    position:relative;
    overflow:hidden;
    padding:32px;
    border:1px solid rgba(80,190,255,.35);
    border-radius:24px;

    background:
    linear-gradient(
        120deg,
        rgba(12,65,105,.95),
        rgba(35,25,85,.95)
    );

    box-shadow:0 15px 55px rgba(0,100,255,.12);
}

.jon-hero h1 {
    font-size:clamp(28px,5vw,46px);
    font-weight:800;
    letter-spacing:-1px;
    color:#e8fbff;
}

.jon-card {
    background:rgba(17,43,72,.78);
    border:1px solid rgba(100,190,255,.22);
    border-radius:18px;
    padding:20px;
    min-height:145px;
    transition:.3s;
}

.jon-card:hover {
    border-color:#39d8ff;
    transform:translateY(-3px);
}

.jon-caption {
    color:#9ccce9;
}

.stButton > button {
    border-radius:12px;
    border:1px solid rgba(70,190,255,.35);
    background:linear-gradient(110deg,#075b91,#3930a0);
    color:white;
    font-weight:700;
    min-height:42px;
}

.stButton > button:hover {
    border-color:#43dcff;
    color:white;
    box-shadow:0 0 20px rgba(0,190,255,.2);
}

[data-testid="stChatMessage"] {
    border:1px solid rgba(100,190,255,.12);
    border-radius:16px;
}

footer {
    visibility:hidden;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 4. API KEY MANAGEMENT
# ============================================================

def get_api_key():

    session_key = st.session_state.get(
        "api_key", ""
    ).strip()

    if session_key:
        return session_key

    try:
        secret_key = st.secrets.get(
            "GEMINI_API_KEY", ""
        )

        if secret_key:
            return str(secret_key).strip()

    except Exception:
        pass

    return ""


def get_client():

    key = get_api_key()

    if not key:
        raise ValueError(
            "API kaliti topilmadi. "
            "Sozlamalar bo'limiga Gemini API kalitini kiriting."
        )

    return genai.Client(api_key=key)


# ============================================================
# 5. AI ERROR HANDLING
# ============================================================

def explain_api_error(error):

    message = str(error)
    lower = message.lower()

    if "429" in lower or "resource_exhausted" in lower:
        return (
            "⚠️ Gemini API kvotasi yoki so'rov limiti tugagan.\n\n"
            "Quyidagilarni tekshiring:\n"
            "- Google AI Studio'dagi API kvotasi.\n"
            "- Billing va tarif holati.\n"
            "- Model uchun mavjud limitlar.\n\n"
            "Kvota tugagan bo'lsa, kutish yoki tarifni "
            "o'zgartirish talab qilinishi mumkin."
        )

    if "403" in lower or "permission_denied" in lower:
        return (
            "API ruxsati rad etildi. API kaliti, loyiha "
            "va modelga kirish huquqini tekshiring."
        )

    if "401" in lower or "unauthenticated" in lower:
        return (
            "API kaliti noto'g'ri yoki yaroqsiz. "
            "Sozlamalardan kalitni qayta kiriting."
        )

    if "404" in lower or "not_found" in lower:
        return (
            "Tanlangan Gemini modeli topilmadi yoki "
            "undan foydalanish imkoniyati yo'q. "
            "Sozlamalardan model nomini tekshiring."
        )

    if "timeout" in lower or "timed out" in lower:
        return (
            "Ulanish vaqti tugadi. Internetni tekshirib, "
            "qayta urinib ko'ring."
        )

    return (
        "Gemini bilan ishlashda xatolik yuz berdi.\n\n"
        + message[:1200]
    )


# ============================================================
# 6. GEMINI AI CORE
# ============================================================

import time
from google import genai

def ask_ai(prompt, extra_context="", history=None):
    client = genai.Client(api_key=get_api_key())

    models = [
        "gemini-3.5-flash",
        "gemini-2.5-flash"
    ]

    full_prompt = f"""
    {JON_CORE}

    {extra_context}

    Foydalanuvchi: {prompt}
    """

    last_error = None

    for model in models:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=full_prompt
                )
                return response.text

            except Exception as e:
                last_error = e
                error_text = str(e)

                if "503" in error_text or "UNAVAILABLE" in error_text:
                    time.sleep(2 * (attempt + 1))
                    continue

                if "429" in error_text:
                    return "Gemini API kvotasi tugagan. Limitni tekshiring."

                break

    return (
        "JON AI vaqtinchalik band. "
        "Iltimos, birozdan keyin qayta urinib ko‘ring."
    )
# ============================================================
# 7. 3D LIVE CHARACTER
# ============================================================

def show_character(height=480, cinematic=False):

    name = json.dumps(
        st.session_state.character_name
    )

    styles = {
        "Hologramma": "0x31eaff",
        "Robot": "0xa8c9df",
        "Virtual inson": "0xffc49a",
        "Multfilm uslubi": "0xffb7e8",
        "Fantastik": "0xb99aff"
    }

    color = styles.get(
        st.session_state.character_style,
        "0x31eaff"
    )

    intro_class = (
        "cinematic" if cinematic else ""
    )

    html = r"""
<!DOCTYPE html>
<html>
<head>

<meta name="viewport"
content="width=device-width,initial-scale=1">

<style>

html,body {
    margin:0;
    padding:0;
    width:100%;
    height:100%;
    overflow:hidden;
    background:#030817;
}

#stage {
    width:100%;
    height:100%;
    position:relative;
    overflow:hidden;
    border-radius:20px;

    background:
    radial-gradient(
        ellipse at center,
        #173f69 0%,
        #101936 45%,
        #030817 100%
    );
}

canvas {
    display:block;
    width:100%;
    height:100%;
}

#label {
    position:absolute;
    bottom:22px;
    width:100%;
    text-align:center;
    color:#b7f8ff;
    font:bold 23px Arial;
    letter-spacing:7px;
    text-shadow:0 0 18px #00dfff;
    pointer-events:none;
}

#subtitle {
    position:absolute;
    bottom:6px;
    width:100%;
    text-align:center;
    color:#82bde5;
    font:11px Arial;
    letter-spacing:3px;
}

#portal {
    position:absolute;
    width:180px;
    height:180px;
    border:2px solid rgba(50,220,255,.7);
    border-radius:50%;
    left:50%;
    top:48%;
    transform:translate(-50%,-50%) scale(.1);
    box-shadow:
    0 0 20px #00dfff,
    inset 0 0 25px #00aaff;
    animation:portalOpen 2s ease-out forwards;
    pointer-events:none;
}

#flash {
    position:absolute;
    inset:0;
    background:#b9faff;
    opacity:0;
    pointer-events:none;
    animation:flashIn 1.2s ease-out;
}

#stars {
    position:absolute;
    inset:0;
    background-image:
    radial-gradient(white 1px,transparent 1px),
    radial-gradient(#42cfff 1px,transparent 1px);
    background-size:55px 55px,90px 90px;
    opacity:.35;
    animation:starsMove 18s linear infinite;
    pointer-events:none;
}

.cinematic #portal {
    animation:portalOpen 2s ease-out forwards;
}

@keyframes portalOpen {
    0% {
        transform:translate(-50%,-50%) scale(.1);
        opacity:0;
    }
    60% {
        opacity:1;
    }
    100% {
        transform:translate(-50%,-50%) scale(2.4);
        opacity:.05;
    }
}

@keyframes flashIn {
    0% {opacity:0}
    15% {opacity:.55}
    100% {opacity:0}
}

@keyframes starsMove {
    from {transform:translateY(0)}
    to {transform:translateY(-80px)}
}

</style>

</head>

<body>

<div id="stage" class="__INTRO__">

    <div id="stars"></div>
    <div id="portal"></div>
    <div id="flash"></div>

    <canvas id="scene"></canvas>

    <div id="label"></div>
    <div id="subtitle">VIRTUAL INTELLIGENCE</div>

</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>

<script>

(() => {

    const stage = document.getElementById("stage");

    const canvas = document.getElementById("scene");

    document.getElementById("label").textContent = __NAME__;

    if (!window.THREE) {
        document.getElementById("subtitle").textContent =
            "3D ENGINE LOADING FAILED";
        return;
    }

    const T = window.THREE;

    const scene = new T.Scene();

    const camera = new T.PerspectiveCamera(
        40,
        1,
        0.1,
        100
    );

    camera.position.z = 7;

    const renderer = new T.WebGLRenderer({
        canvas: canvas,
        antialias: true,
        alpha: true
    });

    renderer.setPixelRatio(
        Math.min(window.devicePixelRatio || 1, 1.5)
    );

    const color = __COLOR__;

    scene.add(
        new T.AmbientLight(0xffffff, 1.5)
    );

    const light = new T.PointLight(
        color,
        3.5,
        25
    );

    light.position.set(3, 3, 5);

    scene.add(light);

    const group = new T.Group();

    scene.add(group);

    const material = new T.MeshStandardMaterial({
        color: color,
        metalness: 0.5,
        roughness: 0.25,
        emissive: color,
        emissiveIntensity: 0.22
    });

    const dark = new T.MeshStandardMaterial({
        color: 0x14243d,
        metalness: 0.6,
        roughness: 0.3
    });

    function sphere(parent, mat, pos, scale) {

        const obj = new T.Mesh(
            new T.SphereGeometry(1, 32, 24),
            mat
        );

        obj.position.set(...pos);

        obj.scale.set(...scale);

        parent.add(obj);

        return obj;
    }

    // Body

    sphere(
        group,
        dark,
        [0, -1, 0],
        [0.85, 1, 0.5]
    );

    // Shoulders

    sphere(
        group,
        material,
        [-0.8, -0.65, 0],
        [0.3, 0.25, 0.3]
    );

    sphere(
        group,
        material,
        [0.8, -0.65, 0],
        [0.3, 0.25, 0.3]
    );

    // Head

    const head = new T.Group();

    head.position.y = 0.65;

    group.add(head);

    sphere(
        head,
        material,
        [0, 0, 0],
        [0.72, 0.85, 0.58]
    );

    const white = new T.MeshBasicMaterial({
        color: 0xffffff
    });

    const black = new T.MeshBasicMaterial({
        color: 0x06101c
    });

    const eyeL = sphere(
        head,
        white,
        [-0.26, 0.08, 0.49],
        [0.14, 0.17, 0.07]
    );

    const eyeR = sphere(
        head,
        white,
        [0.26, 0.08, 0.49],
        [0.14, 0.17, 0.07]
    );

    sphere(
        head,
        black,
        [-0.25, 0.08, 0.55],
        [0.07, 0.09, 0.035]
    );

    sphere(
        head,
        black,
        [0.27, 0.08, 0.55],
        [0.07, 0.09, 0.035]
    );

    const mouth = sphere(
        head,
        black,
        [0, -0.32, 0.53],
        [0.18, 0.035, 0.035]
    );

    // Energy rings

    const ringMaterial = new T.MeshBasicMaterial({
        color: color,
        transparent: true,
        opacity: 0.7
    });

    const ring1 = new T.Mesh(
        new T.TorusGeometry(1.35, 0.018, 8, 100),
        ringMaterial
    );

    ring1.rotation.x = 1.2;

    ring1.position.y = -0.5;

    group.add(ring1);

    const ring2 = new T.Mesh(
        new T.TorusGeometry(1.55, 0.009, 8, 100),
        ringMaterial
    );

    ring2.rotation.x = -0.5;

    group.add(ring2);

    // Particles

    const particleCount = 350;

    const positions = new Float32Array(
        particleCount * 3
    );

    for (let i = 0; i < particleCount; i++) {

        positions[i * 3] = (Math.random() - 0.5) * 7;

        positions[i * 3 + 1] = (Math.random() - 0.5) * 7;

        positions[i * 3 + 2] = (Math.random() - 0.5) * 3;
    }

    const particleGeometry = new T.BufferGeometry();

    particleGeometry.setAttribute(
        "position",
        new T.BufferAttribute(positions, 3)
    );

    const particleMaterial = new T.PointsMaterial({
        color: color,
        size: 0.025,
        transparent: true,
        opacity: 0.7
    });

    const particles = new T.Points(
        particleGeometry,
        particleMaterial
    );

    scene.add(particles);

    const clock = new T.Clock();

    function resize() {

        const w = Math.max(stage.clientWidth, 1);

        const h = Math.max(stage.clientHeight, 1);

        camera.aspect = w / h;

        camera.updateProjectionMatrix();

        renderer.setSize(w, h, false);
    }

    if (window.ResizeObserver) {
        new ResizeObserver(resize).observe(stage);
    } else {
        window.addEventListener("resize", resize);
    }

    resize();

    let animationId;

    function animate() {

        animationId = requestAnimationFrame(animate);

        const t = clock.getElapsedTime();

        group.position.y = Math.sin(t * 1.4) * 0.08;

        group.rotation.y = Math.sin(t * 0.5) * 0.15;

        head.rotation.x = Math.sin(t * 0.8) * 0.04;

        head.rotation.y = Math.sin(t * 0.7) * 0.05;

        ring1.rotation.z += 0.003;

        ring2.rotation.z -= 0.002;

        particles.rotation.y = t * 0.025;

        mouth.scale.y =
            0.035 + Math.abs(Math.sin(t * 3)) * 0.035;

        renderer.render(scene, camera);
    }

    animate();

    window.addEventListener("pagehide", () => {
        cancelAnimationFrame(animationId);
        renderer.dispose();
        particleGeometry.dispose();
    });

})();

</script>

</body>
</html>
"""

    html = (
        html.replace("__NAME__", name)
        .replace("__COLOR__", color)
        .replace("__INTRO__", intro_class)
    )

    components.html(
        html,
        height=height,
        scrolling=False
    )


# ============================================================
# 8. HOME PAGE
# ============================================================

def render_home():

    st.markdown("""
    <div class="jon-hero">
        <h1>💎 JON AI LIVE 8.0</h1>
        <p>Professional Virtual Intelligence System</p>
        <p>Ta'lim · Ijod · Tahlil · AI suhbat · Virtual personaj</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("JON — sizning virtual yordamchingiz")

    show_character(
        height=470,
        cinematic=st.session_state.intro_enabled
    )

    st.markdown(
        "### ✨ Xush kelibsiz!"
    )

    a, b, c = st.columns(3)

    with a:
        st.markdown("""
        <div class="jon-card">
            <h3>🤖 AI Intelligence</h3>
            <p>Gemini AI yordamida aqlli suhbat.</p>
        </div>
        """, unsafe_allow_html=True)

    with b:
        st.markdown("""
        <div class="jon-card">
            <h3>👤 Live Character</h3>
            <p>3D virtual personaj va animatsiyalar.</p>
        </div>
        """, unsafe_allow_html=True)

    with c:
        st.markdown("""
        <div class="jon-card">
            <h3>🚀 Creative Studio</h3>
            <p>Ta'lim va ijodiy vositalar.</p>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    if st.button("🤖 JON bilan suhbatni boshlash"):
        st.session_state.page = PAGES[1]
        st.rerun()


# ============================================================
# 9. AI CHAT
# ============================================================

def render_chat():

    st.title("🤖 JON AI — Chat")

    st.caption(
        "JON AI bilan savol-javob va professional suhbat."
    )

    if st.button("🗑️ Suhbat tarixini tozalash"):
        st.session_state.messages = []
        st.rerun()

    for item in st.session_state.messages:

        with st.chat_message(item["role"]):

            st.markdown(item["content"])

    prompt = st.chat_input(
        "JON AI ga yozing..."
    )

    if prompt:

        history = list(
            st.session_state.messages
        )

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.write(prompt)

        with st.chat_message("assistant"):

            with st.spinner(
                "JON AI fikrlamoqda..."
            ):

                answer = ask_ai(
                    prompt,
                    history
                )

            st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })


# ============================================================
# 10. EDUCATION CENTER
# ============================================================

def render_education():

    st.title("🎓 JON AI — Ta'lim markazi")

    topic = st.selectbox(
        "Yo'nalish",
        [
            "Ingliz tili",
            "IELTS",
            "Multilevel",
            "Grammatika",
            "Lug'at",
            "Matematika",
            "Dasturlash",
            "Boshqa fan"
        ]
    )

    mode = st.selectbox(
        "Yordam turi",
        [
            "Mavzuni tushuntirish",
            "Test tuzish",
            "Javoblarni tekshirish",
            "Reja tuzish",
            "Mashq yaratish",
            "Misollar berish"
        ]
    )
st.selectbox(
    "Bilim darajasi",
    [
        "Boshlang‘ich",
        "O‘rta",
        "Yuqori",
        "Professional"
    ],
    key="education_level"
)

level = st.session_state.education_level

question = st.text_area(
    "Mavzu yoki topshiriqni kiriting",
    placeholder="Masalan: Present Perfect mavzusini tushuntiring..."
)

if st.button("🚀 Darsni boshlash", key="education_start"):
    if not question.strip():
        st.warning("Iltimos, mavzu yoki topshiriq kiriting.")
    else:
        prompt = f"""
        Fan: {topic}
        Yordam turi: {mode}
        Bilim darajasi: {level}
        Topshiriq: {question}

        O‘quvchiga tushunarli, tartibli va sifatli
        ta'limiy material tayyorla.
        Zarur bo‘lsa misollar va mashqlar ham ber.
        """
        with st.spinner("JON AI dars tayyorlamoqda..."):
            answer = ask_ai(prompt)

        st.markdown("### 📚 JON AI javobi")
        st.markdown(answer)


# ============================================================
# 11. POWERPOINT STUDIO
# ============================================================

def render_powerpoint():

    st.title("📊 JON AI — PowerPoint Studio")

    st.write(
        "Sun'iy intellekt yordamida taqdimot rejasini "
        "va PowerPoint faylini tayyorlang."
    )

    topic = st.text_input(
        "Taqdimot mavzusi",
        placeholder="Masalan: Artificial Intelligence"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        min_value=3,
        max_value=20,
        value=8
    )

    language = st.selectbox(
        "Taqdimot tili",
        ["O‘zbekcha", "English", "Русский"]
    )

    style = st.selectbox(
        "Taqdimot uslubi",
        [
            "Professional",
            "Minimalistik",
            "Zamonaviy",
            "Ta'limiy",
            "Biznes"
        ]
    )

    if st.button("✨ Taqdimot yaratish"):

        if not topic.strip():
            st.warning("Taqdimot mavzusini kiriting.")
            return

        prompt = f"""
        Quyidagi mavzuda PowerPoint uchun
        {slide_count} ta slaydlik material tayyorla.

        Mavzu: {topic}
        Til: {language}
        Uslub: {style}

        Har bir slaydni quyidagi formatda yoz:

        SLIDE 1: Sarlavha
        TITLE: ...
        CONTENT:
        - ...
        - ...
        - ...

        Har bir slayd uchun alohida sarlavha
        va 3-5 ta mazmunli punkt yoz.
        """

        with st.spinner("Taqdimot mazmuni yaratilmoqda..."):
            result = ask_ai(prompt)

        st.session_state["ppt_content"] = result

    if st.session_state.get("ppt_content"):

        st.markdown("### 📑 Taqdimot mazmuni")

        st.markdown(st.session_state.ppt_content)

        if st.button("📥 PowerPoint faylini tayyorlash"):

            try:
                presentation = Presentation()

                for line in st.session_state.ppt_content.splitlines():

                    line = line.strip()

                    if not line:
                        continue

                    if line.upper().startswith("SLIDE "):

                        slide = presentation.slides.add_slide(
                            presentation.slide_layouts[1]
                        )

                        slide.shapes.title.text = line

                        body = slide.placeholders[1]
                        body.text = ""

                    elif line.startswith("TITLE:"):

                        if "slide" in locals():
                            slide.shapes.title.text = (
                                line.replace("TITLE:", "").strip()
                            )

                    elif line.startswith("-"):

                        if "slide" in locals():
                            paragraph = body.text_frame.add_paragraph()
                            paragraph.text = line.lstrip("- ").strip()
                            paragraph.level = 0

                if len(presentation.slides) == 0:

                    slide = presentation.slides.add_slide(
                        presentation.slide_layouts[1]
                    )

                    slide.shapes.title.text = topic

                    slide.placeholders[1].text = (
                        st.session_state.ppt_content[:3000]
                    )

                output = BytesIO()

                presentation.save(output)

                output.seek(0)

                st.download_button(
                    "⬇️ PowerPoint faylini yuklab olish",
                    data=output,
                    file_name="JON_AI_Presentation.pptx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    )
                )

            except Exception as error:
                st.error(f"PowerPoint yaratishda xatolik: {error}")


# ============================================================
# 12. TRANSLATOR
# ============================================================

def render_translator():

    st.title("🌐 JON AI — Professional Translator")

    text_to_translate = st.text_area(
        "Tarjima qilinadigan matn",
        height=200,
        placeholder="Matnni shu yerga kiriting..."
    )

    col1, col2 = st.columns(2)

    with col1:
        source_language = st.selectbox(
            "Asl til",
            [
                "Avtomatik aniqlash",
                "O‘zbek",
                "English",
                "Русский",
                "Turkish",
                "Arabic",
                "French",
                "German"
            ]
        )

    with col2:
        target_language = st.selectbox(
            "Tarjima tili",
            [
                "O‘zbek",
                "English",
                "Русский",
                "Turkish",
                "Arabic",
                "French",
                "German"
            ]
        )

    translation_style = st.selectbox(
        "Tarjima uslubi",
        [
            "Tabiiy va ravon",
            "Akademik",
            "Rasmiy",
            "Oddiy suhbat",
            "So‘zma-so‘z"
        ]
    )

    if st.button("🌍 Tarjima qilish"):

        if not text_to_translate.strip():
            st.warning("Tarjima uchun matn kiriting.")
            return

        prompt = f"""
        Quyidagi matnni tarjima qil.

        Asl til: {source_language}
        Maqsad til: {target_language}
        Uslub: {translation_style}

        Talablar:
        - Ma'noni saqla.
        - Grammatik jihatdan to‘g‘ri tarjima qil.
        - Keraksiz izoh qo‘shma.
        - Ismlar va raqamlarni saqla.

        MATN:
        {text_to_translate}
        """

        with st.spinner("Tarjima qilinmoqda..."):
            result = ask_ai(prompt)

        st.markdown("### ✅ Tarjima natijasi")
        st.text_area(
            "Natija",
            value=result,
            height=220
        )

        st.download_button(
            "📥 Tarjimani TXT shaklida yuklab olish",
            data=result,
            file_name="JON_AI_translation.txt",
            mime="text/plain"
        )


# ============================================================
# 13. IMAGE PROMPT STUDIO
# ============================================================

def render_image_prompt():

    st.title("🎨 JON AI — Image Prompt Studio")

    st.write(
        "Professional AI rasm generatorlari uchun "
        "batafsil promptlar tayyorlang."
    )

    subject = st.text_input(
        "Rasm mavzusi",
        placeholder="Masalan: Futuristic city"
    )

    art_style = st.selectbox(
        "San'at uslubi",
        [
            "Realistic",
            "Cinematic",
            "3D Render",
            "Anime",
            "Digital Art",
            "Fantasy",
            "Minimalist",
            "Photorealistic"
        ]
    )

    lighting = st.selectbox(
        "Yoritish",
        [
            "Cinematic lighting",
            "Soft lighting",
            "Neon lighting",
            "Natural lighting",
            "Dramatic lighting"
        ]
    )

    ratio = st.selectbox(
        "Rasm formati",
        [
            "1:1",
            "16:9",
            "9:16",
            "4:3",
            "3:2"
        ]
    )

    details = st.text_area(
        "Qo‘shimcha talablar",
        placeholder="Ranglar, fon, kamera burchagi..."
    )

    if st.button("✨ Professional prompt yaratish"):

        if not subject.strip():
            st.warning("Rasm mavzusini kiriting.")
            return

        prompt = f"""
        Create a professional image generation prompt.

        Subject: {subject}
        Art style: {art_style}
        Lighting: {lighting}
        Aspect ratio: {ratio}
        Additional details: {details}

        Requirements:
        - Highly detailed visual description.
        - Professional composition.
        - Realistic materials and textures where appropriate.
        - High quality rendering.
        - Clear camera and lighting direction.

        Return:
        1. English prompt.
        2. Uzbek explanation.
        3. Negative prompt.
        """

        with st.spinner("Prompt yaratilmoqda..."):
            result = ask_ai(prompt)

        st.markdown("### 🖼️ Tayyor prompt")
        st.markdown(result)

        st.download_button(
            "📥 Promptni yuklab olish",
            data=result,
            file_name="JON_AI_Image_Prompt.txt",
            mime="text/plain"
        )


# ============================================================
# 14. EXCEL ANALYSIS
# ============================================================

def render_excel():

    st.title("📈 JON AI — Excel Analysis")

    st.write(
        "Excel va CSV ma'lumotlarini tahlil qilish."
    )

    uploaded_file = st.file_uploader(
        "Excel yoki CSV faylni yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file:

        try:

            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("Fayl muvaffaqiyatli yuklandi.")

            st.subheader("📋 Ma'lumotlar jadvali")

            st.dataframe(
                df,
                use_container_width=True
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("Qatorlar", len(df))

            with col2:
                st.metric("Ustunlar", len(df.columns))

            with col3:
                st.metric(
                    "Bo‘sh kataklar",
                    int(df.isna().sum().sum())
                )

            st.subheader("📊 Statistik ma'lumotlar")

            st.dataframe(
                df.describe(include="all").transpose(),
                use_container_width=True
            )

            if st.button("🤖 AI orqali tahlil qilish"):

                sample = df.head(30).to_string()

                prompt = f"""
                Quyidagi jadvalni professional tahlil qil.

                Ma'lumotlar:
                {sample}

                Quyidagilarni aniqlashga harakat qil:
                1. Jadvalning umumiy mazmuni.
                2. Muhim tendensiyalar.
                3. Ehtimoliy muammolar.
                4. Qiziqarli kuzatuvlar.
                5. Amaliy tavsiyalar.

                Faqat jadvalda mavjud ma'lumotlarga tayan.
                Yetarli ma'lumot bo‘lmasa, buni ayt.
                """

                with st.spinner("AI ma'lumotlarni tahlil qilmoqda..."):
                    result = ask_ai(prompt)

                st.markdown("### 🧠 AI tahlili")
                st.markdown(result)

            csv_data = df.to_csv(index=False).encode("utf-8-sig")

            st.download_button(
                "📥 CSV faylini yuklab olish",
                data=csv_data,
                file_name="JON_AI_Export.csv",
                mime="text/csv"
            )

        except Exception as error:
            st.error(f"Faylni o‘qishda xatolik: {error}")


# ============================================================
# 15. TASK MANAGER
# ============================================================

def render_tasks():

    st.title("📝 JON AI — Vazifalar boshqaruvi")

    st.write(
        "Kundalik vazifalaringizni rejalashtiring."
    )

    with st.form("new_task_form"):

        task_text = st.text_input(
            "Yangi vazifa",
            placeholder="Masalan: JON AI loyihasini tekshirish"
        )

        priority = st.selectbox(
            "Muhimlik darajasi",
            ["Oddiy", "Muhim", "Juda muhim"]
        )

        submitted = st.form_submit_button(
            "➕ Vazifa qo‘shish"
        )

    if submitted:

        if task_text.strip():

            st.session_state.tasks.append({
                "id": str(time.time_ns()),
                "text": task_text.strip(),
                "priority": priority,
                "done": False
            })

            st.success("Vazifa qo‘shildi.")
            st.rerun()

        else:
            st.warning("Vazifa matnini kiriting.")

    st.divider()

    if not st.session_state.tasks:

        st.info("Hozircha vazifalar mavjud emas.")

    else:

        completed = sum(
            1 for task in st.session_state.tasks
            if task["done"]
        )

        total = len(st.session_state.tasks)

        st.progress(
            completed / total
        )

        st.write(
            f"Bajarilgan: {completed} / {total}"
        )

        for task in list(st.session_state.tasks):

            col1, col2, col3 = st.columns([6, 2, 1])

            with col1:

                checked = st.checkbox(
                    task["text"],
                    value=task["done"],
                    key=f"task_{task['id']}"
                )

                if checked != task["done"]:

                    task["done"] = checked

                    st.rerun()

            with col2:
                st.caption(task["priority"])

            with col3:

                if st.button(
                    "🗑️",
                    key=f"delete_{task['id']}"
                ):

                    st.session_state.tasks = [
                        item
                        for item in st.session_state.tasks
                        if item["id"] != task["id"]
                    ]

                    st.rerun()

        if st.button("🧹 Barcha bajarilgan vazifalarni o‘chirish"):

            st.session_state.tasks = [
                task for task in st.session_state.tasks
                if not task["done"]
            ]

            st.rerun()


# ============================================================
# 16. CAMERA
# ============================================================

def render_camera():

    st.title("📷 JON AI — Camera Studio")

    st.write(
        "Kamera orqali suratga oling va uni JON AI "
        "bilan tahlil qilish uchun tayyorlang."
    )

    photo = st.camera_input(
        "Kamerani ochish va suratga olish"
    )

    if photo:

        st.image(
            photo,
            caption="Olingan surat",
            use_container_width=True
        )

        st.success(
            "Surat muvaffaqiyatli olindi."
        )

        st.info(
            "Hozirgi versiyada kamera suratga olishni "
            "qo‘llab-quvvatlaydi. Suratni AI orqali "
            "tahlil qilish funksiyasi keyingi bosqichda "
            "alohida ulanadi."
        )

        st.download_button(
            "📥 Suratni yuklab olish",
            data=photo.getvalue(),
            file_name="JON_AI_Photo.jpg",
            mime="image/jpeg"
        )


# ============================================================
# 17. CHARACTER STUDIO
# ============================================================

def render_character():

    st.title("👤 JON AI — Live Character Studio")

    st.write(
        "JON virtual personajining ko‘rinishini sozlang."
    )

    name = st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    style_options = [
        "Hologramma",
        "Robot",
        "Virtual inson",
        "Multfilm uslubi",
        "Fantastik"
    ]

    current_style = st.session_state.character_style

    selected_index = (
        style_options.index(current_style)
        if current_style in style_options
        else 0
    )

    style = st.selectbox(
        "Personaj uslubi",
        style_options,
        index=selected_index
    )

    if st.button("💾 Sozlamalarni saqlash"):

        st.session_state.character_name = (
            name.strip() or "JON"
        )

        st.session_state.character_style = style

        st.success("Personaj sozlamalari saqlandi.")

        st.rerun()

    st.divider()

    st.subheader("✨ Jonli personaj")

    show_character(
        height=520,
        cinematic=False
    )

    st.caption(
        "Bu versiyada geometrik 3D personaj mavjud. "
        "Haqiqiy inson mimikasi, ovoz va lab sinxronizatsiyasi "
        "keyingi ishlab chiqish bosqichlarida qo‘shiladi."
    )


# ============================================================
# 18. SETTINGS
# ============================================================

def render_settings():

    st.title("⚙️ JON AI — Sozlamalar")

    st.subheader("🔐 Gemini API")

    st.warning(
        "API kalitingizni hech kimga yubormang va "
        "GitHub kodiga joylashtirmang."
    )

    current_key = st.session_state.get(
        "api_key", ""
    )

    api_key = st.text_input(
        "Gemini API kaliti",
        value=current_key,
        type="password",
        placeholder="Gemini API key"
    )

    if st.button("💾 API kalitni saqlash"):

        if api_key.strip():

            st.session_state.api_key = api_key.strip()

            st.success(
                "API kalit joriy sessiya uchun saqlandi."
            )

        else:

            st.warning(
                "Iltimos, API kalitni kiriting."
            )

    if st.button("🗑️ Kiritilgan API kalitni tozalash"):

        st.session_state.api_key = ""

        st.success("Sessiyadagi API kalit tozalandi.")

        st.rerun()

    st.divider()

    st.subheader("🧠 AI modeli")

    model_options = [
        DEFAULT_MODEL,
        "gemini-2.5-flash",
        "gemini-2.5-pro"
    ]

    current_model = st.session_state.get(
        "model", DEFAULT_MODEL
    )

    if current_model not in model_options:
        model_options.insert(0, current_model)

    selected_model = st.selectbox(
        "Modelni tanlang",
        model_options
    )

    if st.button("Modelni saqlash"):

        st.session_state.model = selected_model

        st.success("Model tanlovi saqlandi.")

    st.divider()

    st.subheader("🎬 Kirish animatsiyasi")

    intro = st.toggle(
        "Cinematic intro",
        value=st.session_state.intro_enabled
    )

    st.session_state.intro_enabled = intro

    st.divider()

    st.subheader("📱 Ilova haqida")

    st.write(f"JON AI LIVE versiyasi: {APP_VERSION}")

    st.write(
        "JON AI — sun'iy intellekt asosidagi "
        "virtual yordamchi loyihasi."
    )
# ============================================================
# 19. PREMIUM
# ============================================================

def render_premium():

    st.title("💎 JON AI — Premium")

    st.markdown("""
    <div class="jon-hero">
        <h1>JON AI PREMIUM</h1>
        <p>Professional AI Creative & Intelligence Studio</p>
        <p>Kelajak texnologiyalari bitta platformada.</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("🚀 Rejalashtirilgan Premium imkoniyatlar")

    features = [
        (
            "🎨 AI Image Studio",
            "Professional AI rasm yaratish va tahrirlash."
        ),
        (
            "🎬 AI Video Studio",
            "AI yordamida video yaratish va montaj."
        ),
        (
            "🎵 AI Music Studio",
            "Original musiqa va fon musiqalarini yaratish."
        ),
        (
            "👤 AI Character Animation",
            "3D personajlar, mimika va animatsiyalar."
        ),
        (
            "🎞️ Movie Maker",
            "Ssenariy, sahna, ovoz va musiqa bilan video loyihalar."
        ),
        (
            "🏢 Professional Branding",
            "Kompaniya logotiplari va brend dizayni."
        ),
        (
            "🖼️ Photo to Movie",
            "Rasmlardan videolar yaratish."
        ),
        (
            "🎭 Photo to Cartoon",
            "Rasmlarni multfilm uslubiga o‘zgartirish."
        ),
        (
            "🎙️ Voice Assistant",
            "Ovozli suhbat va nutq sintezini qo‘llash."
        )
    ]

    for title, description in features:

        st.markdown(
            f"""
            <div class="jon-card" style="margin-bottom:12px;">
                <h3>{title}</h3>
                <p>{description}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.info(
        "Diqqat: yuqoridagi Premium imkoniyatlar "
        "ushbu kodda hozircha reja sifatida ko‘rsatilgan. "
        "Ularni ishlatish uchun tegishli AI xizmatlari, "
        "API va media modullari alohida ulanadi."
    )


# ============================================================
# 20. MAIN NAVIGATION
# ============================================================

def main():

    with st.sidebar:

        st.markdown(
            """
            <h1 style="
                color:#5deaff;
                text-align:center;
                letter-spacing:3px;
            ">💎 JON AI</h1>
            <p style="
                color:#9ccce9;
                text-align:center;
            ">LIVE INTELLIGENCE</p>
            """,
            unsafe_allow_html=True
        )

        st.divider()

        selected_page = st.radio(
            "Navigatsiya",
            PAGES,
            index=PAGES.index(
                st.session_state.page
            ),
            key="main_navigation"
        )

        st.session_state.page = selected_page

        st.divider()

        st.caption(
            f"JON AI LIVE v{APP_VERSION}"
        )

        st.caption(
            "Professional Virtual Intelligence"
        )

    page = st.session_state.page

    if page == PAGES[0]:

        render_home()

    elif page == PAGES[1]:

        render_chat()

    elif page == PAGES[2]:

        render_education()

    elif page == PAGES[3]:

        render_powerpoint()

    elif page == PAGES[4]:

        render_translator()

    elif page == PAGES[5]:

        render_image_prompt()

    elif page == PAGES[6]:

        render_excel()

    elif page == PAGES[7]:

        render_tasks()

    elif page == PAGES[8]:

        render_camera()

    elif page == PAGES[9]:

        render_character()

    elif page == PAGES[10]:

        render_settings()

    elif page == PAGES[11]:

        render_premium()


# ============================================================
# 21. START APPLICATION
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        st.error(
            "JON AI ishga tushishida xatolik yuz berdi."
        )

        st.exception(error)
