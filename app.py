# ============================================================
# JON AI LIVE 6.0 — YAXLIT BIR FAYLLI VERSIYA
# Streamlit + Gemini AI + Three.js 3D Character
# ============================================================

import json
import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from pptx import Presentation


# ============================================================
# 1. ASOSIY SOZLAMALAR
# ============================================================

st.set_page_config(
    page_title="JON AI LIVE",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "6.0"
MODEL = "gemini-2.5-flash"

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
    "💎 Premium",
]

JON_CORE = """
Sen JON AI nomli professional virtual yordamchisan.

Asosiy til — o‘zbek tili.
Foydalanuvchi boshqa tilda yozsa, o‘sha tilda javob ber.
Ravon, aniq, hurmatli va amaliy javob ber.
Ma’lumotni to‘qib chiqarma.
Ishonching komil bo‘lmasa, buni ochiq ayt.
Murakkab vazifalarni bosqichma-bosqich tushuntir.
Kod yozganda to‘liq va ishlatish mumkin bo‘lgan kod ber.
Maxfiy kalitlar va parollarni so‘rama yoki oshkor qilma.
"""

DEFAULTS = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "api_key": "",
    "api_input": "",
    "api_status": "",
    "selected_page": PAGES[0],
    "last_translation": "",
    "last_prompt": "",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# 2. DIZAYN
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: Inter, sans-serif;
    }

    .stApp {
        background:
        radial-gradient(
            ellipse at 10% 0%,
            rgba(22,85,130,.22),
            transparent 38%
        ),
        linear-gradient(
            145deg,
            #07111f 0%,
            #0a1425 52%,
            #101329 100%
        );
        color: #edf5ff;
    }

    [data-testid="stHeader"] {
        background: rgba(0,0,0,0);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #0b1729,
            #10152b
        );
        border-right: 1px solid rgba(88,190,255,.18);
    }

    .block-container {
        padding-top: 1.6rem;
        padding-bottom: 2.5rem;
        max-width: 1400px;
    }

    .jon-hero {
        padding: 30px;
        border-radius: 24px;
        border: 1px solid rgba(83,205,255,.25);
        background: linear-gradient(
            120deg,
            rgba(20,69,111,.62),
            rgba(38,31,83,.45)
        );
        box-shadow: 0 12px 45px rgba(0,0,0,.18);
    }

    .jon-hero h1 {
        font-size: clamp(30px,5vw,52px);
        margin: 0;
        color: #eafaff;
    }

    .jon-hero p {
        color: #a9c6e5;
        font-size: 16px;
        margin-top: 10px;
    }

    .jon-card {
        border: 1px solid rgba(110,190,255,.18);
        background: rgba(17,35,59,.58);
        padding: 18px;
        border-radius: 18px;
        min-height: 120px;
    }

    .jon-muted {
        color: #9db5d0;
    }

    div.stButton > button {
        border-radius: 12px;
        border: 1px solid rgba(95,195,255,.28);
        min-height: 42px;
        font-weight: 600;
    }

    div[data-testid="stChatMessage"] {
        border: 1px solid rgba(100,170,230,.12);
        border-radius: 15px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 3. GEMINI API
# ============================================================

def get_api_key():
    key = str(
        st.session_state.get("api_key", "")
    ).strip()

    if key:
        return key

    try:
        return str(
            st.secrets.get("GEMINI_API_KEY", "")
        ).strip()
    except Exception:
        return ""


def get_client():
    key = get_api_key()

    if not key:
        raise RuntimeError(
            "Gemini API kaliti topilmadi. "
            "Sozlamalar bo‘limida kalitni kiriting."
        )

    return genai.Client(api_key=key)


def ask_ai(prompt, extra_context="", history=None):
    """
    Gemini AI chaqiruvi.
    Oxirgi 12 ta xabar suhbat konteksti sifatida uzatiladi.
    """

    try:
        client = get_client()

        history_lines = []

        for item in (history or [])[-12:]:
            role = (
                "Foydalanuvchi"
                if item.get("role") == "user"
                else "JON AI"
            )

            content = str(
                item.get("content", "")
            )[:12000]

            history_lines.append(
                f"{role}: {content}"
            )

        full_prompt = f"""
{JON_CORE}

Qo‘shimcha kontekst:
{extra_context}

Oldingi suhbat:
{chr(10).join(history_lines)}

Yangi topshiriq:
{prompt}
"""

        response = client.models.generate_content(
            model=MODEL,
            contents=full_prompt,
        )

        answer = getattr(response, "text", None)

        if answer:
            return answer.strip()

        return "JON AI javob qaytara olmadi."

    except Exception as exc:
        return f"AI bilan ishlashda xatolik: {exc}"


def test_api_connection():
    try:
        client = get_client()

        response = client.models.generate_content(
            model=MODEL,
            contents="Faqat OK deb javob ber.",
        )

        if getattr(response, "text", None):
            st.session_state.api_status = (
                "🟢 Gemini API muvaffaqiyatli ulandi."
            )
        else:
            st.session_state.api_status = (
                "⚠️ API javobi bo‘sh qaytdi."
            )

    except Exception as exc:
        st.session_state.api_status = (
            f"🔴 Ulanish xatosi: {exc}"
        )


# ============================================================
# 4. 3D JONLI PERSONAJ
# ============================================================

def show_character(speaking=False, height=540):

    name = json.dumps(
        str(st.session_state.character_name)
    )

    style = st.session_state.character_style

    style_map = {
        "Hologramma": "hologram",
        "Robot": "robot",
        "Virtual inson": "human",
        "Multfilm uslubi": "cartoon",
        "Fantastik": "fantasy",
    }

    visual_style = style_map.get(
        style,
        "hologram"
    )

    speaking_js = (
        "true" if speaking else "false"
    )

    html = r"""
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport"
content="width=device-width,initial-scale=1">

<style>

html,body {
    margin:0;
    width:100%;
    height:100%;
    overflow:hidden;
    background:#07101d;
    font-family:Arial;
}

#stage {
    width:100%;
    height:100%;
    position:relative;
    overflow:hidden;
    border-radius:20px;
    background:
    radial-gradient(
        ellipse at 50% 38%,
        #19395d 0%,
        #0b172b 58%,
        #050a13 100%
    );
}

canvas {
    width:100%;
    height:100%;
    display:block;
}

.hud {
    position:absolute;
    top:15px;
    left:18px;
    right:18px;
    display:flex;
    justify-content:space-between;
    color:#9cecff;
    font-size:11px;
    letter-spacing:2px;
    pointer-events:none;
}

#name {
    position:absolute;
    bottom:38px;
    width:100%;
    text-align:center;
    color:#e1fbff;
    font-size:23px;
    font-weight:bold;
    letter-spacing:7px;
    text-shadow:0 0 15px #00dfff;
}

#status {
    position:absolute;
    bottom:13px;
    width:100%;
    text-align:center;
    color:#75e8ff;
    font-size:10px;
    letter-spacing:3px;
}

</style>
</head>

<body>

<div id="stage">

    <div class="hud">
        <span>JON AI LIVE • 3D ENGINE</span>
        <span>● ONLINE</span>
    </div>

    <canvas id="scene"></canvas>

    <div id="name"></div>

    <div id="status">
        VIRTUAL INTELLIGENCE
    </div>

</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>

<script>

(()=>{

const stage = document.getElementById('stage');
const canvas = document.getElementById('scene');

const nameNode = document.getElementById('name');
const status = document.getElementById('status');

nameNode.textContent = __NAME__;

const style = '__STYLE__';
const speaking = __SPEAKING__;

if(!window.THREE) {
    status.textContent = '3D ENGINE YUKLANMADI';
    return;
}

const T = window.THREE;

const scene = new T.Scene();

const camera = new T.PerspectiveCamera(
    38,
    stage.clientWidth / stage.clientHeight,
    0.1,
    100
);

camera.position.set(0,0.1,8.2);

const renderer = new T.WebGLRenderer({
    canvas,
    antialias:true,
    alpha:true
});

renderer.setPixelRatio(
    Math.min(window.devicePixelRatio || 1,1.7)
);

renderer.setSize(
    stage.clientWidth,
    stage.clientHeight
);

renderer.outputEncoding = T.sRGBEncoding;

const colors = {
    hologram:0x31eaff,
    robot:0xa8c9df,
    human:0xffc49a,
    cartoon:0xffb7e8,
    fantasy:0xb99aff
};

const accent = colors[style] || colors.hologram;

scene.add(
    new T.AmbientLight(0x9bcfff,1.25)
);

const light = new T.PointLight(
    accent,
    2.8,
    20
);

light.position.set(3,3,5);
scene.add(light);

const fill = new T.PointLight(
    0x287bff,
    1.6,
    16
);

fill.position.set(-3,-1,3);
scene.add(fill);

const character = new T.Group();
scene.add(character);

const mat = (
    color,
    metal=.35,
    rough=.3,
    emissive=0
) => new T.MeshStandardMaterial({
    color,
    metalness:metal,
    roughness:rough,
    emissive,
    emissiveIntensity:
        style === 'hologram' ? .4 : .12
});

const skinColor =
    style === 'human' ? 0xe9ad89 :
    style === 'cartoon' ? 0xffc4dc :
    style === 'fantasy' ? 0xb9a0ff :
    style === 'robot' ? 0x9eb7c9 :
    0x55dff4;

const skin = mat(skinColor,.4,.3,accent);
const dark = mat(0x101b34,.7,.25);

const glow = new T.MeshBasicMaterial({
    color:accent
});

const white = new T.MeshStandardMaterial({
    color:0xffffff
});

const black = new T.MeshBasicMaterial({
    color:0x07101d
});

function sphere(parent,material,pos,scale) {

    const m = new T.Mesh(
        new T.SphereGeometry(1,32,24),
        material
    );

    m.position.set(...pos);
    m.scale.set(...scale);

    parent.add(m);

    return m;
}


// TANA

sphere(
    character,
    dark,
    [0,-1.45,0],
    [.92,1.05,.48]
);

sphere(
    character,
    skin,
    [0,-.78,.02],
    [.86,.42,.48]
);

sphere(
    character,
    dark,
    [-.82,-.95,0],
    [.28,.72,.32]
);

sphere(
    character,
    dark,
    [.82,-.95,0],
    [.28,.72,.32]
);

sphere(
    character,
    skin,
    [0,-.28,0],
    [.28,.38,.3]
);


// BOSH

const head = new T.Group();

head.position.set(0,.62,0);

character.add(head);

sphere(
    head,
    skin,
    [0,0,0],
    [.72,.88,.58]
);

sphere(
    head,
    dark,
    [0,.55,0],
    [.76,.25,.61]
);

sphere(
    head,
    glow,
    [0,.77,0],
    [.18,.07,.12]
);

sphere(
    head,
    dark,
    [-.72,-.02,0],
    [.13,.25,.16]
);

sphere(
    head,
    dark,
    [.72,-.02,0],
    [.13,.25,.16]
);


// KO‘ZLAR

const le = sphere(
    head,
    white,
    [-.27,.08,.49],
    [.16,.19,.075]
);

const re = sphere(
    head,
    white,
    [.27,.08,.49],
    [.16,.19,.075]
);

sphere(
    head,
    black,
    [-.25,.07,.555],
    [.075,.10,.035]
);

sphere(
    head,
    black,
    [.29,.07,.555],
    [.075,.10,.035]
);


// BURUN

sphere(
    head,
    skin,
    [0,-.12,.55],
    [.075,.12,.08]
);


// OG‘IZ

const mouth = sphere(
    head,
    black,
    [0,-.34,.535],
    [.19,.035,.035]
);


// KO‘K NUR

sphere(
    character,
    glow,
    [0,-1.25,.48],
    [.16,.16,.06]
);


// HALQALAR

const ringMat = new T.MeshBasicMaterial({
    color:accent,
    transparent:true,
    opacity:.6
});

const ring1 = new T.Mesh(
    new T.TorusGeometry(1.42,.012,8,100),
    ringMat
);

ring1.rotation.x = 1.12;
ring1.position.y = -.55;

character.add(ring1);


const ring2 = new T.Mesh(
    new T.TorusGeometry(1.65,.008,8,100),
    ringMat
);

ring2.rotation.x = 1.35;
ring2.rotation.y = .55;
ring2.position.y = -.55;

character.add(ring2);


// YULDUZLAR

const positions = [];

for(let i=0;i<140;i++) {

    positions.push(
        (Math.random()-.5)*9,
        (Math.random()-.5)*6,
        (Math.random()-.5)*4-1
    );

}

const geo = new T.BufferGeometry();

geo.setAttribute(
    'position',
    new T.Float32BufferAttribute(positions,3)
);

const stars = new T.Points(
    geo,
    new T.PointsMaterial({
        color:accent,
        size:.025,
        transparent:true,
        opacity:.8
    })
);

scene.add(stars);


// ANIMATSIYA

let blinkAt = 2;
let blinkUntil = 0;

const clock = new T.Clock();

function resize() {

    const w = stage.clientWidth;
    const h = stage.clientHeight;

    camera.aspect = w/h;
    camera.updateProjectionMatrix();

    renderer.setSize(w,h);

}

window.addEventListener('resize',resize);


function animate() {

    requestAnimationFrame(animate);

    const t = clock.getElapsedTime();

    character.position.y =
        Math.sin(t*1.35)*.07;

    character.rotation.y =
        Math.sin(t*.55)*.13;

    head.rotation.y =
        Math.sin(t*.8)*.08;

    head.rotation.x =
        Math.sin(t*.65)*.035;

    ring1.rotation.z += .003;
    ring2.rotation.z -= .002;

    stars.rotation.y += .0005;

    if(t > blinkAt) {

        blinkUntil = t+.14;

        blinkAt = t+3.2+Math.random()*2.2;

    }

    const blink = t < blinkUntil;

    le.scale.y = blink ? .025 : .19;
    re.scale.y = blink ? .025 : .19;

    if(speaking) {

        mouth.scale.y =
            .035+Math.abs(Math.sin(t*12))*.12;

        mouth.scale.x =
            .19+Math.abs(Math.sin(t*12))*.06;

        status.textContent = 'JON AI GAPIRMOQDA';

    } else {

        mouth.scale.y = .035;
        mouth.scale.x = .19;

        status.textContent = 'VIRTUAL INTELLIGENCE';

    }

    renderer.render(scene,camera);

}

animate();

})();

</script>
</body>
</html>
"""

    html = (
        html
        .replace("__NAME__", name)
        .replace("__STYLE__", visual_style)
        .replace("__SPEAKING__", speaking_js)
    )

    components.html(
        html,
        height=height,
        scrolling=False
    )


# ============================================================
# 5. POWERPOINT YARATISH
# ============================================================

def parse_slides(text):

    cleaned = re.sub(
        r"```(?:json)?",
        "",
        str(text),
        flags=re.I
    ).replace("```", "").strip()

    start = cleaned.find("[")
    end = cleaned.rfind("]") + 1

    if start < 0 or end <= start:
        raise ValueError(
            "AI javobida JSON slaydlar massivi topilmadi."
        )

    data = json.loads(
        cleaned[start:end]
    )

    if not isinstance(data,list) or not data:
        raise ValueError(
            "Slaydlar ro‘yxati bo‘sh yoki noto‘g‘ri."
        )

    return data


def create_presentation(topic,slides_data):

    prs = Presentation()

    for item in slides_data:

        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )

        slide.shapes.title.text = str(
            item.get("title",topic)
        )

        body = slide.placeholders[1].text_frame

        body.clear()

        points = item.get("points",[])

        if isinstance(points,str):
            points = [points]

        for idx,point in enumerate(points):

            paragraph = (
                body.paragraphs[0]
                if idx == 0
                else body.add_paragraph()
            )

            paragraph.text = str(point)
            paragraph.level = 0

    output = BytesIO()

    prs.save(output)

    output.seek(0)

    return output


# ============================================================
# 6. BOSH SAHIFA
# ============================================================

def render_home():

    st.markdown(
        """
        <div class="jon-hero">

            <h1>💎 JON AI LIVE</h1>

            <p>
            Virtual Intelligence System ·
            Professional AI Assistant
            </p>

            <p>
            Ta’lim, ijod, tahlil va kundalik
            vazifalar uchun yagona AI muhiti.
            </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    a,b,c = st.columns(3)

    with a:

        st.markdown(
            """
            <div class="jon-card">
                <h3>🤖 AI Intelligence</h3>
                <p class="jon-muted">
                Gemini asosidagi aqlli suhbat va yordam.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    with b:

        st.markdown(
            """
            <div class="jon-card">
                <h3>👤 Live Character</h3>
                <p class="jon-muted">
                Three.js asosidagi jonli 3D personaj.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c:

        st.markdown(
            """
            <div class="jon-card">
                <h3>🚀 Creative Tools</h3>
                <p class="jon-muted">
                Hujjat, tarjima, ta’lim va tahlil vositalari.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.subheader("JON personaji")

    show_character(height=430)

    st.caption(
        "3D personaj animatsiyasi brauzerda ishlaydi. "
        "Bu prototip fotorealistik inson modeli emas."
    )


# ============================================================
# 7. AI CHAT
# ============================================================

def render_chat():

    st.title("🤖 JON AI — Aqlli suhbat")

    for item in st.session_state.messages:

        with st.chat_message(item["role"]):
            st.markdown(item["content"])

    prompt = st.chat_input(
        "JON AI ga savolingizni yozing..."
    )

    if prompt:

        st.session_state.messages.append({
            "role":"user",
            "content":prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):

            with st.spinner("JON AI fikrlamoqda..."):

                answer = ask_ai(
                    prompt,
                    history=st.session_state.messages[:-1]
                )

            st.markdown(answer)

        st.session_state.messages.append({
            "role":"assistant",
            "content":answer
        })


# ============================================================
# 8. TA’LIM MARKAZI
# ============================================================

def render_education():

    st.title("🎓 JON AI — Ta’lim markazi")

    topic = st.selectbox(
        "Yo‘nalish",
        [
            "Ingliz tili",
            "IELTS",
            "Multilevel",
            "Grammatika",
            "Lug‘at",
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
            "Mashq yaratish"
        ]
    )

    level = st.selectbox(
        "Daraja",
        [
            "Boshlang‘ich",
            "O‘rta",
            "Yuqori"
        ]
    )

    question = st.text_area(
        "Mavzu yoki topshiriq",
        placeholder="Masalan: Present Perfect mavzusi"
