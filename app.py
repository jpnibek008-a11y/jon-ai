
# ============================================================
# JON AI LIVE 10.0
# NEXT GENERATION VIRTUAL INTELLIGENCE
# UNIFIED STREAMLIT APPLICATION
# ============================================================

import json
import re
import time
import random
import base64
import html
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt

# ============================================================
# 1. APPLICATION CONFIGURATION
# ============================================================

APP_NAME = "JON AI LIVE"
APP_VERSION = "10.0"

# Model nomi Settings orqali o'zgartirilishi mumkin.
DEFAULT_MODEL = "gemini-2.5-flash"

st.set_page_config(
    page_title=f"{APP_NAME} {APP_VERSION}",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 2. APPLICATION NAVIGATION
# ============================================================

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

# ============================================================
# 3. AI SYSTEM PERSONALITY
# ============================================================

JON_CORE = """
Siz JON AI nomli professional sun'iy intellekt yordamchisiz.

Sizning vazifalaringiz:

1. Foydalanuvchiga o'zbek tilida sifatli javob berish.
2. Ingliz, rus va boshqa tillarda yordam berish.
3. Dasturlash va texnik muammolarni yechish.
4. Ta'lim va imtihonlarga tayyorlash.
5. Biznes va startap g'oyalarini ishlab chiqish.
6. Tarjima va matn yozishda yordam berish.
7. Excel va ma'lumotlarni tahlil qilish.
8. Taqdimotlar tayyorlash.
9. Ijodiy loyihalarda yordam berish.

QOIDALAR:

- Bilmagan ma'lumotni to'qib chiqarmang.
- Noaniq ma'lumotni aniq deb ko'rsatmang.
- Javoblarni tartibli va tushunarli yozing.
- Murakkab vazifalarni bosqichlarga ajrating.
- Professional va hurmatli ohangda muloqot qiling.
- Foydalanuvchining maxfiyligini hurmat qiling.
"""

# ============================================================
# 4. SESSION STATE
# ============================================================

DEFAULTS = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "character_animation": "Raqs",
    "api_key": "",
    "model": DEFAULT_MODEL,
    "intro_enabled": True,
    "copied_prompt": "",
    "last_answer": "",
    "motivation_quote": "",
    "premium": False,
    "selected_page": "🏠 Bosh sahifa",
    "user_name": "",
    "camera_enabled": False
}

for key, value in DEFAULTS.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# 5. MOTIVATIONAL CONTENT
# ============================================================

MOTIVATIONAL_QUOTES = [

    (
        "TO‘XTAMA. SEN HALI ORZU QILGAN MANZILINGGA YETMAGANSAN!",
        "Charchashing mumkin. Ammo taslim bo‘lma."
    ),

    (
        "BUGUNGI HARAKATING — ERTANGI G‘ALABANG.",
        "Kichik qadamlar katta natijalarga olib keladi."
    ),

    (
        "QO‘RQMA. BOSHLAGAN ISHINGNI OXIRIGA YETKAZ!",
        "Mukammal bo‘lish shart emas. Muhimi davom etish."
    ),

    (
        "SEN O‘YLAGANINGDAN HAM KUCHLISAN!",
        "Qiyinchiliklar seni chiniqtirish uchun mavjud."
    ),

    (
        "BUGUN TASLIM BO‘LMA!",
        "Ertangi o‘zing bugungi harakating uchun rahmat aytadi."
    ),

    (
        "ORZULAR KUTMAYDI. HARAKATNI BUGUN BOSHLА!",
        "O‘z tarixingni o‘zing yarat."
    )

]

if not st.session_state.motivation_quote:

    st.session_state.motivation_quote = random.choice(
        MOTIVATIONAL_QUOTES
    )

# ============================================================
# 6. GEMINI API KEY MANAGEMENT
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

        raise ValueError(
            "Gemini API kalitini Sozlamalar bo'limiga kiriting."
        )

    return genai.Client(api_key=key)


# ============================================================
# 7. ERROR HANDLING
# ============================================================

def explain_api_error(error):

    message = str(error)
    lower = message.lower()

    if "429" in message or "resource_exhausted" in lower:

        return (
            "API kvotasi yoki so'rov limiti tugagan. "
            "Google AI Studio limitlarini tekshiring."
        )

    if "503" in message or "unavailable" in lower:

        return (
            "Gemini serveri vaqtincha band. "
            "Keyinroq qayta urinib ko'ring."
        )

    if "401" in message or "403" in message:

        return (
            "API kalit noto'g'ri yoki ruxsat yetarli emas."
        )

    if "404" in message or "not found" in lower:

        return (
            "Model topilmadi. Sozlamalardan "
            "mavjud Gemini modelini tanlang."
        )

    if "timeout" in lower:

        return "Ulanish vaqti tugadi. Internetni tekshiring."

    return f"Gemini xatosi: {message[:400]}"


# ============================================================
# 8. GEMINI AI ENGINE
# ============================================================

def ask_ai(prompt, extra_context="", history=None):

    try:

        client = get_client()

        model = st.session_state.get(
            "model",
            DEFAULT_MODEL
        )

        conversation = [JON_CORE]

        for item in (history or [])[-12:]:

            role = item.get("role", "user")
            content = item.get("content", "")

            if not content:
                continue

            role_name = (
                "Foydalanuvchi"
                if role == "user"
                else "JON AI"
            )

            conversation.append(
                f"{role_name}: {content}"
            )

        if extra_context:

            conversation.append(
                f"Qo'shimcha kontekst:\n{extra_context}"
            )

        conversation.append(
            f"Foydalanuvchi: {prompt}"
        )

        conversation.append("JON AI:")

        response = client.models.generate_content(
            model=model,
            contents="\n\n".join(conversation)
        )

        answer = getattr(response, "text", None)

        if not answer:

            return (
                "AI javob qaytarmadi. "
                "Savolingizni qayta yuboring."
            )

        return answer.strip()

    except Exception as error:

        return explain_api_error(error)


# ============================================================
# 9. MODEL CONNECTION TEST
# ============================================================

def check_model():

    try:

        client = get_client()

        model = st.session_state.get(
            "model",
            DEFAULT_MODEL
        )

        response = client.models.generate_content(
            model=model,
            contents="Salom. Faqat OK deb javob ber."
        )

        return bool(
            getattr(response, "text", None)
        )

    except Exception as error:

        st.error(explain_api_error(error))

        return False


# ============================================================
# 10. DESIGN SYSTEM
# ============================================================

st.markdown("""
<style>

@import url(
'https://fonts.googleapis.com/css2?family=Orbitron:wght@500;700;800&family=Inter:wght@300;400;500;600;700&display=swap'
);

:root {
    --blue: #26baff;
    --purple: #9b5cff;
    --background: #050817;
}

.stApp {

    background:
        radial-gradient(
            ellipse at 15% 10%,
            #17265a 0%,
            transparent 35%
        ),
        radial-gradient(
            ellipse at 85% 20%,
            #25124d 0%,
            transparent 35%
        ),
        linear-gradient(
            145deg,
            #050817,
            #080d20 60%,
            #09051c
        );

    color: #edf4ff;

    font-family: Inter, sans-serif;

}

h1, h2, h3 {

    color: #edf4ff !important;

}

.hero-title {

    font-family: Orbitron, sans-serif;

    font-weight: 800;

    letter-spacing: 4px;

    text-align: center;

    font-size: clamp(28px, 5vw, 58px);

    background: linear-gradient(
        90deg,
        #ffffff,
        #43caff,
        #a879ff
    );

    -webkit-background-clip: text;

    -webkit-text-fill-color: transparent;

}

.hero-subtitle {

    color: #9bb8df;

    text-align: center;

    letter-spacing: 4px;

    font-size: 11px;

}

.panel {

    background: linear-gradient(
        145deg,
        #111c3bdb,
        #10102bdd
    );

    border: 1px solid #4e8cff44;

    border-radius: 18px;

    padding: 20px;

    margin-bottom: 15px;

    box-shadow: 0 0 25px #176dff12;

}

.feature-card {

    background: linear-gradient(
        145deg,
        #111c38,
        #15102e
    );

    border: 1px solid #438dff44;

    border-radius: 15px;

    padding: 18px;

    min-height: 115px;

}

.motivation-frame {

    position: relative;

    overflow: hidden;

    padding: 22px;

    margin: 20px auto;

    text-align: center;

    border: 1px solid #42cfff;

    border-radius: 18px;

    background: linear-gradient(
        125deg,
        #091c3a,
        #201141
    );

    box-shadow: 0 0 22px #26baff33;

    animation: motivationGlow 4s infinite;

}

.motivation-label {

    color: #65dcff;

    font-size: 11px;

    font-weight: 800;

    letter-spacing: 3px;

}

.motivation-quote {

    color: white;

    font-family: Orbitron, sans-serif;

    font-size: clamp(15px, 2.2vw, 23px);

    line-height: 1.6;

    font-weight: 800;

}

.motivation-subtext {

    color: #b9cce8;

    font-size: 13px;

    margin-top: 10px;

}

@keyframes motivationGlow {

    0%,100% {

        box-shadow: 0 0 12px #26baff22;

    }

    50% {

        box-shadow: 0 0 28px #26baff66;

    }

}

div.stButton > button {

    border-radius: 12px;

    border: 1px solid #438dff77;

    background: linear-gradient(
        100deg,
        #123c75,
        #39206b
    );

    color: white;

    font-weight: 600;

    min-height: 42px;

    transition: 0.3s;

}

div.stButton > button:hover {

    border-color: #54d5ff;

    color: white;

    box-shadow: 0 0 15px #36baff55;

}

[data-testid="stSidebar"] {

    background: linear-gradient(
        180deg,
        #080e23,
        #100d29
    );

    border-right: 1px solid #3475b944;

}

[data-testid="stChatMessage"] {

    background: #101a35;

    border: 1px solid #3d6eb533;

    border-radius: 14px;

}

.stTextInput input,
.stTextArea textarea {

    background: #0d1730 !important;

    color: white !important;

    border-color: #3475b966 !important;

}

hr {

    border-color: #3c69a633;

}

@media(max-width:600px) {

    .motivation-frame {

        padding: 16px;

    }

    .hero-title {

        letter-spacing: 2px;

    }

}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 1-QISM TUGADI
# KEYINGI QISM:
# HUMAN + ANDROID 3D ENGINE
# ============================================================

# ============================================================
# JON AI LIVE 10.0
# 2-QISM — BIR DONA MUSTAQIL AI PERSONAJ
# ============================================================

def detect_jon_mood():
    """Oxirgi AI javobidan personaj kayfiyatini aniqlash."""

    answer = str(
        st.session_state.get("last_answer", "")
    ).lower()

    if any(word in answer for word in [
        "tabrik", "ajoyib", "zo'r", "muvaffaqiyat",
        "excellent", "congratulations", "great"
    ]):
        return "happy"

    if any(word in answer for word in [
        "afsus", "xato", "xavf", "ehtiyot",
        "sorry", "warning", "error"
    ]):
        return "serious"

    if any(word in answer for word in [
        "salom", "xush kelibsiz", "qandaysiz",
        "hello", "welcome", "hi"
    ]):
        return "greeting"

    if any(word in answer for word in [
        "o'ylab", "tahlil", "sababi", "demak",
        "analysis", "because", "therefore"
    ]):
        return "thinking"

    return "neutral"


def show_character(height=640, cinematic=True):
    """
    JON — yagona markaziy AI personaj.
    Mustaqil animatsiya, mimika va maxsus effektlar.
    """

    character_name = st.session_state.get(
        "character_name", "JON"
    )

    character_style = st.session_state.get(
        "character_style", "Hologramma"
    )

    animation_mode = st.session_state.get(
        "animation", "Tabiiy harakat"
    )

    mood = detect_jon_mood()

    style_map = {
        "Hologramma": "hologram",
        "Realistik robot": "robot",
        "Multfilm uslubi": "cartoon",
        "Fantastik": "fantasy"
    }

    visual_style = style_map.get(
        character_style, "hologram"
    )

    html_doc = r"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport"
            content="width=device-width, initial-scale=1">

      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }

        body {
          overflow: hidden;
          background: transparent;
          font-family: Arial, sans-serif;
        }

        #jon-stage {
          width: 100%;
          height: HEIGHTpx;
          position: relative;
          overflow: hidden;
          border-radius: 24px;
          background:
            radial-gradient(
              ellipse at 50% 65%,
              #123b68 0%,
              #071a36 42%,
              #020817 100%
            );
          border: 1px solid rgba(0,220,255,.35);
          box-shadow:
            0 0 35px rgba(0,180,255,.2),
            inset 0 0 45px rgba(0,180,255,.08);
        }

        #jon-canvas {
          position: absolute;
          inset: 0;
          width: 100%;
          height: 100%;
        }

        .jon-heading {
          position: absolute;
          top: 20px;
          left: 0;
          width: 100%;
          text-align: center;
          color: #78eaff;
          font-size: 13px;
          letter-spacing: 5px;
          text-shadow: 0 0 15px #00dfff;
          pointer-events: none;
        }

        .jon-name {
          position: absolute;
          bottom: 47px;
          width: 100%;
          text-align: center;
          color: #d9faff;
          font-size: 17px;
          font-weight: bold;
          letter-spacing: 7px;
          text-shadow: 0 0 15px #00dfff;
          pointer-events: none;
        }

        .jon-status {
          position: absolute;
          bottom: 21px;
          width: 100%;
          text-align: center;
          color: #6dffdc;
          font-size: 10px;
          letter-spacing: 4px;
          pointer-events: none;
        }

        .jon-dot {
          display: inline-block;
          width: 7px;
          height: 7px;
          border-radius: 50%;
          background: #00ffc3;
          margin-right: 8px;
          box-shadow: 0 0 12px #00ffc3;
          animation: blink 1.2s infinite;
        }

        @keyframes blink {
          50% { opacity: .2; }
        }

        .jon-corner {
          position: absolute;
          width: 35px;
          height: 35px;
          border-color: rgba(0,220,255,.5);
          pointer-events: none;
        }

        .top-left {
          top: 12px;
          left: 12px;
          border-top: 1px solid;
          border-left: 1px solid;
        }

        .top-right {
          top: 12px;
          right: 12px;
          border-top: 1px solid;
          border-right: 1px solid;
        }

        .bottom-left {
          bottom: 12px;
          left: 12px;
          border-bottom: 1px solid;
          border-left: 1px solid;
        }

        .bottom-right {
          bottom: 12px;
          right: 12px;
          border-bottom: 1px solid;
          border-right: 1px solid;
        }
      </style>
    </head>

    <body>
      <div id="jon-stage">
        <div class="jon-heading">
          JON AI • AUTONOMOUS INTELLIGENCE
        </div>

        <div class="jon-corner top-left"></div>
        <div class="jon-corner top-right"></div>
        <div class="jon-corner bottom-left"></div>
        <div class="jon-corner bottom-right"></div>

        <canvas id="jon-canvas"></canvas>

        <div class="jon-name">JON</div>

        <div class="jon-status">
          <span class="jon-dot"></span>
          ONLINE • READY TO INTERACT
        </div>
      </div>

      <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>

      <script>
      (() => {
        const stage = document.getElementById("jon-stage");
        const canvas = document.getElementById("jon-canvas");

        if (!window.THREE) {
          stage.innerHTML +=
            '<div style="color:white;text-align:center;padding:200px 15px;">3D tizim yuklanmadi. Internetni tekshiring.</div>';
          return;
        }

        const THREE = window.THREE;

        const STYLE = VISUAL_STYLE;
        const MODE = ANIMATION_MODE;
        const MOOD = CHARACTER_MOOD;

        const scene = new THREE.Scene();
        scene.fog = new THREE.FogExp2(0x020817, 0.035);

        const camera = new THREE.PerspectiveCamera(
          38, 1, 0.1, 100
        );

        camera.position.set(0, 2.2, 9.5);
        camera.lookAt(0, 2.1, 0);

        const renderer = new THREE.WebGLRenderer({
          canvas: canvas,
          alpha: true,
          antialias: true
        });

        renderer.setPixelRatio(
          Math.min(window.devicePixelRatio || 1, 2)
        );

        renderer.shadowMap.enabled = true;

        // YORUG'LIK
        scene.add(
          new THREE.AmbientLight(0x9bcfff, 1.25)
        );

        const keyLight = new THREE.PointLight(
          0x00cfff, 4, 18
        );

        keyLight.position.set(0, 5, 4);
        scene.add(keyLight);

        const sideLight = new THREE.PointLight(
          0x714dff, 3, 15
        );

        sideLight.position.set(-4, 2, -2);
        scene.add(sideLight);

        const backLight = new THREE.PointLight(
          0x00ffcc, 2.5, 15
        );

        backLight.position.set(4, 3, -3);
        scene.add(backLight);

        // MATERIALLAR
        const isHologram = STYLE === "hologram";
        const isCartoon = STYLE === "cartoon";
        const isFantasy = STYLE === "fantasy";

        const bodyMaterial = new THREE.MeshStandardMaterial({
          color: isCartoon ? 0x55bce9 : 0x5b9bc4,
          metalness: isHologram ? 0.72 : 0.48,
          roughness: 0.3,
          emissive: 0x063b68,
          emissiveIntensity: 0.8,
          transparent: isHologram,
          opacity: isHologram ? 0.86 : 1
        });

        const darkMaterial = new THREE.MeshStandardMaterial({
          color: 0x10253d,
          metalness: 0.78,
          roughness: 0.25
        });

        const faceMaterial = new THREE.MeshStandardMaterial({
          color: isCartoon ? 0xffc69c : 0x8cb9d8,
          metalness: 0.25,
          roughness: 0.45,
          transparent: isHologram,
          opacity: isHologram ? 0.9 : 1
        });

        const glowColor = isFantasy ? 0xd08aff : 0x35edff;

        const eyeMaterial = new THREE.MeshBasicMaterial({
          color: glowColor
        });

        // YORDAMCHI FUNKSIYALAR
        function sphere(parent, material,
                        x, y, z, sx, sy, sz) {
          const object = new THREE.Mesh(
            new THREE.SphereGeometry(1, 32, 24),
            material
          );

          object.position.set(x, y, z);
          object.scale.set(sx, sy, sz);
          parent.add(object);

          return object;
        }

        function cylinder(parent, material,
                          x, y, z, rt, rb, h) {
          const object = new THREE.Mesh(
            new THREE.CylinderGeometry(rt, rb, h, 24),
            material
          );

          object.position.set(x, y, z);
          parent.add(object);

          return object;
        }

        // JONNING YAGONA TANASI
        const jon = new THREE.Group();
        scene.add(jon);

        const torso = new THREE.Group();
        jon.add(torso);

        // OYOQLAR
        const leftLeg = new THREE.Group();
        leftLeg.position.set(-0.22, 1.35, 0);
        jon.add(leftLeg);

        cylinder(
          leftLeg, darkMaterial,
          0, -0.25, 0, 0.17, 0.13, 0.9
        );

        sphere(
          leftLeg, bodyMaterial,
          0, -0.72, 0.08, 0.22, 0.11, 0.34
        );

        const rightLeg = new THREE.Group();
        rightLeg.position.set(0.22, 1.35, 0);
        jon.add(rightLeg);

        cylinder(
          rightLeg, darkMaterial,
          0, -0.25, 0, 0.17, 0.13, 0.9
        );

        sphere(
          rightLeg, bodyMaterial,
          0, -0.72, 0.08, 0.22, 0.11, 0.34
        );

        // TANA
        sphere(
          torso, bodyMaterial,
          0, 2.35, 0, 0.62, 0.77, 0.36
        );

        // KO'KRAK ENERGIYA MARKAZI
        const chestCore = sphere(
          torso, eyeMaterial,
          0, 2.48, 0.34, 0.13, 0.13, 0.055
        );

        const chestRing = new THREE.Mesh(
          new THREE.TorusGeometry(0.22, 0.012, 8, 48),
          new THREE.MeshBasicMaterial({
            color: glowColor,
            transparent: true,
            opacity: 0.75
          })
        );

        chestRing.position.set(0, 2.48, 0.35);
        torso.add(chestRing);

        // BO'YIN
        cylinder(
          torso, darkMaterial,
          0, 3.02, 0, 0.17, 0.17, 0.28
        );

        // BOSH — ALOHIDA HARAKATLANADI
        const head = new THREE.Group();
        head.position.set(0, 3.58, 0);
        jon.add(head);

        sphere(
          head, faceMaterial,
          0, 0, 0, 0.48, 0.55, 0.4
        );

        // Soch / yuqori qoplama
        if (!isHologram && !STYLE === "robot") {
          sphere(
            head, darkMaterial,
            0, 0.28, -0.04,
            0.47, 0.25, 0.39
          );
        }

        // Peshona energiya belgisi
        sphere(
          head, eyeMaterial,
          0, 0.27, 0.385,
          0.055, 0.055, 0.025
        );

        // KO'ZLAR
        const leftEye = sphere(
          head, eyeMaterial,
          -0.17, 0.04, 0.36,
          0.055, 0.035, 0.025
        );

        const rightEye = sphere(
          head, eyeMaterial,
          0.17, 0.04, 0.36,
          0.055, 0.035, 0.025
        );

        // KO'Z ATROFIDAGI YORUG'LIK
        sphere(
          head, eyeMaterial,
          -0.17, 0.04, 0.34,
          0.11, 0.075, 0.018
        ).material = eyeMaterial.clone();

        sphere(
          head, eyeMaterial,
          0.17, 0.04, 0.34,
          0.11, 0.075, 0.018
        ).material = eyeMaterial.clone();

        // OG'IZ — GAPIRISH ANIMATSIYASI
        const mouth = new THREE.Mesh(
          new THREE.SphereGeometry(1, 20, 12),
          eyeMaterial
        );

        mouth.position.set(0, -0.2, 0.385);
        mouth.scale.set(0.09, 0.018, 0.025);
        head.add(mouth);

        // QULOQ / SENSORLAR
        sphere(
          head, darkMaterial,
          -0.47, -0.02, 0,
          0.09, 0.18, 0.13
        );

        sphere(
          head, darkMaterial,
          0.47, -0.02, 0,
          0.09, 0.18, 0.13
        );

        sphere(
          head, eyeMaterial,
          -0.5, -0.02, 0.04,
          0.045, 0.045, 0.045
        );

        sphere(
          head, eyeMaterial,
          0.5, -0.02, 0.04,
          0.045, 0.045, 0.045
        );

        // QO'LLAR
        const leftArm = new THREE.Group();
        leftArm.position.set(-0.58, 2.77, 0);
        jon.add(leftArm);

        sphere(
          leftArm, bodyMaterial,
          0, -0.12, 0, 0.2, 0.22, 0.2
        );

        cylinder(
          leftArm, bodyMaterial,
          -0.08, -0.48, 0,
          0.15, 0.12, 0.62
        );

        sphere(
          leftArm, darkMaterial,
          -0.08, -0.82, 0,
          0.15, 0.15, 0.15
        );

        const rightArm = new THREE.Group();
        rightArm.position.set(0.58, 2.77, 0);
        jon.add(rightArm);

        sphere(
          rightArm, bodyMaterial,
          0, -0.12, 0, 0.2, 0.22, 0.2
        );

        cylinder(
          rightArm, bodyMaterial,
          0.08, -0.48, 0,
          0.15, 0.12, 0.62
        );

        sphere(
          rightArm, darkMaterial,
          0.08, -0.82, 0,
          0.15, 0.15, 0.15
        );

        // ORQA HALQA
        const halo = new THREE.Mesh(
          new THREE.TorusGeometry(0.75, 0.018, 8, 100),
          new THREE.MeshBasicMaterial({
            color: glowColor,
            transparent: true,
            opacity: 0.8
          })
        );

        halo.position.set(0, 4.35, -0.15);
        halo.rotation.x = 0.3;
        jon.add(halo);

        // ENERGIYA HALQALARI
        const energyRings = [];

        for (let i = 0; i < 3; i++) {
          const ring = new THREE.Mesh(
            new THREE.TorusGeometry(
              1.1 + i * 0.38,
              0.012,
              8,
              100
            ),
            new THREE.MeshBasicMaterial({
              color: i % 2 ? 0x8c5aff : 0x00dfff,
              transparent: true,
              opacity: 0.35
            })
          );

          ring.position.y = 0.1 + i * 0.08;
          ring.rotation.x = Math.PI / 2;
          ring.scale.set(1.2, 1.2, 1);
          scene.add(ring);

          energyRings.push(ring);
        }

        // YERDAGI ENERGIYA MAYDONI
        const platform = new THREE.Mesh(
          new THREE.CircleGeometry(1.7, 80),
          new THREE.MeshBasicMaterial({
            color: 0x00aaff,
            transparent: true,
            opacity: 0.12
          })
        );

        platform.rotation.x = -Math.PI / 2;
        platform.position.y = 0.02;
        scene.add(platform);

        // ZARRACHALAR
        const particleCount = 1000;
        const particlePositions = new Float32Array(
          particleCount * 3
        );

        const particleSpeeds = [];

        for (let i = 0; i < particleCount; i++) {
          particlePositions[i * 3] =
            (Math.random() - 0.5) * 10;

          particlePositions[i * 3 + 1] =
            Math.random() * 7;

          particlePositions[i * 3 + 2] =
            (Math.random() - 0.5) * 5;

          particleSpeeds.push(
            0.004 + Math.random() * 0.012
          );
        }

        const particleGeometry = new THREE.BufferGeometry();

        particleGeometry.setAttribute(
          "position",
          new THREE.BufferAttribute(
            particlePositions, 3
          )
        );

        const particleMaterial = new THREE.PointsMaterial({
          color: glowColor,
          size: 0.035,
          transparent: true,
          opacity: 0.85,
          blending: THREE.AdditiveBlending
        });

        const particles = new THREE.Points(
          particleGeometry,
          particleMaterial
        );

        scene.add(particles);

        // MAYDA YULDUZLAR
        const stars = [];

        for (let i = 0; i < 30; i++) {
          const star = new THREE.Mesh(
            new THREE.SphereGeometry(0.025, 8, 8),
            new THREE.MeshBasicMaterial({
              color: 0x9cefff
            })
          );

          star.position.set(
            (Math.random() - 0.5) * 9,
            Math.random() * 6,
            -1.5
          );

          scene.add(star);
          stars.push(star);
        }

        // PAYDO BO'LISH EFFEKTI
        jon.scale.setScalar(0.01);

        const clock = new THREE.Clock();
        let elapsed = 0;
        let blinkTimer = 0;
        let isBlinking = false;

        function resize() {
          const width = stage.clientWidth;
          const height = stage.clientHeight;

          camera.aspect = width / height;
          camera.position.z = width < 500 ? 11 : 9.5;
          camera.updateProjectionMatrix();

          renderer.setSize(width, height, false);
        }

        resize();

        if (window.ResizeObserver) {
          const observer = new ResizeObserver(resize);
          observer.observe(stage);
        } else {
          window.addEventListener("resize", resize);
        }

        // MUSTAQIL HARAKAT
        function animate() {
          requestAnimationFrame(animate);

          const delta = Math.min(
            clock.getDelta(), 0.05
          );

          elapsed += delta;

          const t = elapsed;

          // Personajning paydo bo'lishi
          const entrance = Math.min(t / 2.2, 1);
          const smoothEntrance =
            entrance * entrance * (3 - 2 * entrance);

          const targetScale =
            0.85 * smoothEntrance;

          jon.scale.setScalar(targetScale);

          // Nafas olish va tabiiy harakat
          jon.position.y =
            0.06 * Math.sin(t * 1.8);

          torso.scale.y =
            1 + Math.sin(t * 1.8) * 0.012;

          // Boshning mustaqil harakati
          head.rotation.y =
            Math.sin(t * 0.65) * 0.16;

          head.rotation.z =
            Math.sin(t * 0.9) * 0.035;

          head.rotation.x =
            Math.sin(t * 0.7) * 0.025;

          // Qo'l harakatlari
          if (MODE === "Raqs") {
            leftArm.rotation.z =
              Math.sin(t * 2.5) * 0.55;

            rightArm.rotation.z =
              -Math.sin(t * 2.5) * 0.55;

            jon.rotation.z =
              Math.sin(t * 1.8) * 0.045;

          } else if (MODE === "Energetik") {
            leftArm.rotation.z =
              Math.sin(t * 3.5) * 0.75;

            rightArm.rotation.z =
              -Math.sin(t * 3.5) * 0.75;

          } else if (MODE === "Sokin") {
            leftArm.rotation.z = 0;
            rightArm.rotation.z = 0;

          } else {
            leftArm.rotation.z =
              Math.sin(t * 1.2) * 0.12;

            rightArm.rotation.z =
              -Math.sin(t * 1.2) * 0.12;
          }

          // Kayfiyatga qarab harakat
          if (MOOD === "greeting") {
            rightArm.rotation.z =
              -0.7 + Math.sin(t * 4) * 0.18;
          }

          if (MOOD === "happy") {
            jon.rotation.z =
              Math.sin(t * 2.2) * 0.06;

            leftArm.rotation.z =
              Math.sin(t * 2.5) * 0.35;

            rightArm.rotation.z =
              -Math.sin(t * 2.5) * 0.35;
          }

          if (MOOD === "thinking") {
            head.rotation.z =
              -0.13 + Math.sin(t) * 0.025;

            rightArm.rotation.z = -0.25;
          }

          if (MOOD === "serious") {
            head.rotation.x = -0.04;
          }

          // Ko'z qisish
          blinkTimer += delta;

          if (blinkTimer > 3.5 + Math.random() * 2) {
            blinkTimer = 0;
            isBlinking = true;
          }

          if (isBlinking) {
            leftEye.scale.y = 0.005;
            rightEye.scale.y = 0.005;

            if (blinkTimer > 0.15) {
              isBlinking = false;
              leftEye.scale.y = 0.035;
              rightEye.scale.y = 0.035;
            }
          }

          // Gapirayotgandek og'iz animatsiyasi
    
# ============================================================
# 3-QISM — JON AI BOSH SAHIFASI VA NAVIGATSIYA
# ============================================================


def go_to_page(page):
    st.session_state["selected_page"] = page
    st.rerun()


def render_navigation():
    """JON AI yon navigatsiya paneli."""

    with st.sidebar:

        st.markdown(
            """
            <div style="
                text-align:center;
                padding:15px 5px;
                border-bottom:1px solid #16456b;
                margin-bottom:15px;
            ">
                <div style="
                    font-size:38px;
                    filter:drop-shadow(0 0 12px #00dfff);
                ">💎</div>

                <div style="
                    color:#65eaff;
                    font-size:22px;
                    font-weight:800;
                    letter-spacing:3px;
                ">JON AI</div>

                <div style="
                    color:#8ca9c5;
                    font-size:10px;
                    letter-spacing:3px;
                ">LIVE INTELLIGENCE 10.0</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("### Navigatsiya")

        selected = st.session_state.get(
            "selected_page",
            PAGES[0]
        )

        if selected not in PAGES:
            selected = PAGES[0]

        page = st.radio(
            "Bo‘limni tanlang",
            PAGES,
            index=PAGES.index(selected),
            key="jon_navigation"
        )

        st.session_state["selected_page"] = page

        st.divider()

        st.markdown("### 👤 Personaj sozlamalari")

        st.selectbox(
            "JON ko‘rinishi",
            [
                "Hologramma",
                "Realistik robot",
                "Multfilm uslubi",
                "Fantastik"
            ],
            key="character_style"
        )

        st.selectbox(
            "Harakat uslubi",
            [
                "Raqs",
                "Tabiiy harakat",
                "Sokin",
                "Energetik"
            ],
            key="character_animation"
        )

        st.caption("💎 JON AI LIVE 10.0")
        st.caption("Status: ONLINE")

    return page
    def render_home():
    """JON AI asosiy bosh sahifasi."""

    st.markdown(
        """
        <style>
        .jon-welcome {
            text-align:center;
            padding:8px 5px 12px;
        }

        .jon-welcome h1 {
            font-size:clamp(27px, 5vw, 43px);
            color:#8beaff;
            letter-spacing:3px;
            text-shadow:0 0 20px rgba(0,210,255,.55);
            margin-bottom:5px;
        }

        .jon-welcome p {
            color:#a9c4df;
            font-size:14px;
            letter-spacing:1px;
        }

        .jon-feature {
            min-height:125px;
            padding:18px 12px;
            border-radius:17px;
            text-align:center;
            border:1px solid rgba(0,210,255,.25);
            background:linear-gradient(
                145deg,
                rgba(12,42,76,.92),
                rgba(5,17,37,.96)
            );
            box-shadow:0 0 18px rgba(0,150,255,.08);
            margin-bottom:10px;
        }

        .jon-feature-icon {
            font-size:30px;
            margin-bottom:8px;
        }

        .jon-feature-title {
            color:#8beaff;
            font-size:14px;
            font-weight:700;
        }

        .jon-feature-description {
            color:#9db5ce;
            font-size:11px;
            margin-top:6px;
        }

        .jon-home-caption {
            text-align:center;
            color:#8daac6;
            font-size:12px;
            padding:8px;
        }
        </style>

        <div class="jon-welcome">
            <h1>JON AI</h1>
            <p>
                SIZNING SHAXSIY SUN'IY INTELLEKT HAMKORINGIZ
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    # JONNING AVTOMATIK PAYDO BO'LISHI
    show_character(height=590)

    st.markdown(
        """
        <div class="jon-home-caption">
            JON sizni kutmoqda. Suhbatni boshlang!
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### 🚀 JON AI imkoniyatlari")

    features = [
        {
            "icon": "🤖",
            "title": "AI Chat",
            "description": "JON bilan aqlli suhbat"
        },
        {
            "icon": "🎓",
            "title": "Ta'lim markazi",
            "description": "Bilim va imtihon tayyorgarligi"
        },
        {
            "icon": "📊",
            "title": "PowerPoint",
            "description": "Professional taqdimotlar"
        },
        {
            "icon": "🌐",
            "title": "Tarjimon",
            "description": "Ko‘p tilli tarjima"
        },
        {
            "icon": "🎨",
            "title": "Rasm prompti",
            "description": "Ijodiy tasvir g‘oyalari"
        },
        {
            "icon": "📈",
            "title": "Excel tahlil",
            "description": "Ma'lumotlarni tahlil qilish"
        }
    ]

    for start in range(0, len(features), 3):
        cols = st.columns(3)

        for col, feature in zip(
            cols, features[start:start + 3]
        ):
            with col:
                st.markdown(
                    f"""
                    <div class="jon-feature">
                        <div class="jon-feature-icon">
                            {feature["icon"]}
                        </div>
                        <div class="jon-feature-title">
                            {feature["title"]}
                        </div>
                        <div class="jon-feature-description">
                            {feature["description"]}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if st.button(
                    feature["title"],
                    key="home_" + feature["title"],
                    use_container_width=True
                ):
                    go_to_page(feature["title"])

    st.divider()

    st.markdown("### 💬 JON bilan suhbatni boshlang")

    if st.button(
        "🤖 JON AI CHAT OCHISH",
        use_container_width=True,
        type="primary"
    ):
        go_to_page("AI Chat")

    st.markdown(
        """
        <div class="jon-home-caption">
            JON AI LIVE 10.0 • BUILT FOR THE FUTURE
        </div>
        """,
        unsafe_allow_html=True
    )


def render_character_page():
    """Personajni alohida ko‘rish va sozlash sahifasi."""

    st.title("👤 JON — Autonomous AI Character")

    st.write(
        "JON — JON AI platformasining yagona "
        "markaziy virtual personaji."
    )

    show_character(height=650)

    st.markdown("### ⚙️ Personaj sozlamalari")

    st.text_input(
        "Personaj nomi",
        key="character_name"
    )

    st.caption(
        "Uslub va animatsiya turlarini chap tomondagi "
        "navigatsiya panelidan o‘zgartirishingiz mumkin."
    )


# 3-QISM TUGADI
# KEYINGI QISM: AI CHAT VA JONNING MUSTAQIL QAROR TIZIMI

# ============================================================
# JON AI LIVE 10.0 — 4-QISM
# AI CHAT + GEMINI PERSONAJ KAYFIYATI + MOTIVATSIYA RAMKASI
# ============================================================

import json
import re
from datetime import date


# ------------------------------------------------------------
# 1. GEMINI JAVOBI VA PERSONAJ KAYFIYATINI BIRGA ANIQLASH
# ------------------------------------------------------------

def ask_jon_with_behavior(user_prompt):
    """Gemini bir so‘rovda javob va JON kayfiyatini qaytaradi."""

    try:
        client = get_client()

        if client is None:
            return "Gemini API kalitini Sozlamalar bo‘limiga kiriting."

        model_name = st.session_state.get("model", DEFAULT_MODEL)

        history = st.session_state.get("messages", [])
        history_text = ""

        for item in history[-8:]:
            role = item.get("role", "user")
            content = item.get("content", "")
            history_text += f"\n{role}: {content}"

        system_prompt = f"""
{JON_CORE}

Sen JON nomli aqlli, samimiy va professional AI personajisan.

Foydalanuvchiga aniq, foydali va tushunarli javob ber.
Bilmagan ma'lumotingni to'qib chiqarmagin.

Har bir javob bilan birga personajning kayfiyatini ham tanla.

Faqat quyidagi JSON formatida javob ber:

{{
  "answer": "Foydalanuvchiga beriladigan javob",
  "mood": "neutral",
  "action": "speaking"
}}

mood qiymatlari:
- greeting
- happy
- serious
- thinking
- neutral

action qiymatlari:
- greeting
- speaking
- thinking
- celebrating
- listening

Kayfiyatni suhbat mazmuniga mos tanla.
JSON dan tashqari izoh yozma.
"""

        full_prompt = f"""
{system_prompt}

OLDINGI SUHBAT:
{history_text}

FOYDALANUVCHI SAVOLI:
{user_prompt}
"""

        response = client.models.generate_content(
            model=model_name,
            contents=full_prompt
        )

        raw_text = response.text.strip()

        # JSON blokini ajratib olish
        raw_text = re.sub(
            r"^```(?:json)?\s*|\s*```$",
            "",
            raw_text,
            flags=re.IGNORECASE
        ).strip()

        data = json.loads(raw_text)

        answer = str(data.get("answer", "")).strip()
        mood = str(data.get("mood", "neutral")).lower()
        action = str(data.get("action", "speaking")).lower()

        allowed_moods = {
            "greeting", "happy", "serious",
            "thinking", "neutral"
        }

        allowed_actions = {
            "greeting", "speaking", "thinking",
            "celebrating", "listening"
        }

        if mood not in allowed_moods:
            mood = "neutral"

        if action not in allowed_actions:
            action = "speaking"

        if not answer:
            answer = "Kechirasiz, javobni qayta tayyorlashga urinib ko‘ring."

        st.session_state["jon_mood"] = mood
        st.session_state["jon_action"] = action
        st.session_state["last_answer"] = answer

        return answer

    except json.JSONDecodeError:
        # Model JSON o‘rniga oddiy matn qaytarsa
        try:
            fallback = raw_text.strip()

            if fallback:
                st.session_state["jon_mood"] = "neutral"
                st.session_state["jon_action"] = "speaking"
                st.session_state["last_answer"] = fallback
                return fallback

        except Exception:
            pass

        return "JON javobni qayta tayyorlay olmadi. Yana bir marta urinib ko‘ring."

    except Exception as error:
        return explain_api_error(error)


# ------------------------------------------------------------
# 2. PERSONAJ KAYFIYATINI ANIQLASH
# ------------------------------------------------------------

def detect_jon_mood():
    mood = st.session_state.get("jon_mood", "neutral")

    allowed = {
        "greeting", "happy", "serious",
        "thinking", "neutral"
    }

    if mood in allowed:
        return mood

    return "neutral"


# ------------------------------------------------------------
# 3. JON AI CHAT SAHIFASI
# ------------------------------------------------------------

def render_chat():

    st.title("🤖 JON AI — Intelligent Chat")

    st.caption(
        "JON bilan suhbatlashing. "
        "U javob mazmuniga qarab o‘z kayfiyatini o‘zgartiradi."
    )

    # Oldingi suhbatlarni ko‘rsatish
    messages = st.session_state.get("messages", [])

    for message in messages:
        role = message.get("role", "assistant")
        content = message.get("content", "")

        if role not in ("user", "assistant"):
            continue

        with st.chat_message(role):
            st.markdown(content)

    # Yangi xabar
    user_prompt = st.chat_input(
        "JONga xabaringizni yozing..."
    )

    if user_prompt:

        # Foydalanuvchi xabarini saqlash
        messages = st.session_state.get("messages", [])

        messages.append({
            "role": "user",
            "content": user_prompt
        })

        st.session_state["messages"] = messages

        with st.chat_message("user"):
            st.markdown(user_prompt)

        # JON javobi
        with st.chat_message("assistant"):

            with st.spinner("JON sizni tinglamoqda..."):

                answer = ask_jon_with_behavior(user_prompt)

            st.markdown(answer)

        messages = st.session_state.get("messages", [])

        messages.append({
            "role": "assistant",
            "content": answer
        })

        st.session_state["messages"] = messages

        st.rerun()


# ------------------------------------------------------------
# 4. KUNLIK MOTIVATSIYA RAMKASI
# ------------------------------------------------------------

JON_MOTIVATION_QUOTES = [
    "Buyuk maqsadlar kichik, ammo qat'iy qadamlardan boshlanadi.",
    "Bugungi mehnating ertangi muvaffaqiyatingni yaratadi.",
    "O'zingga ishon. Sen o'ylaganingdan ham kuchlisan.",
    "Mukammallikni kutma. Boshlagin va rivojlan.",
    "Har bir qiyinchilik seni kuchliroq qiladi.",
    "Maqsadingni unutma. Har kuni unga bir qadam yaqinlash.",
    "Muvaffaqiyat tasodif emas, u intizom natijasidir.",
    "Bugun taslim bo'lmasang, ertaga faxrlanasan.",
    "Orzular rejasiz amalga oshmaydi. Harakatni boshlagin.",
    "Kelajaging bugun qilayotgan tanlovlaringdan boshlanadi."
]


def render_motivation_frame():

    today = date.today()
    quote_index = today.toordinal() % len(JON_MOTIVATION_QUOTES)
    quote = JON_MOTIVATION_QUOTES[quote_index]

    st.markdown(
        f"""
        <div style="
            position:relative;
            overflow:hidden;
            padding:28px 22px;
            margin:22px 0;
            border-radius:22px;
            background:
                radial-gradient(
                    circle at 10% 20%,
                    rgba(0,220,255,0.20),
                    transparent 45%
                ),
                linear-gradient(
                    135deg,
                    #07152b,
                    #10102d,
                    #17102b
                );
            border:1px solid rgba(0,220,255,0.55);
            box-shadow:
                0 0 15px rgba(0,220,255,0.20),
                inset 0 0 25px rgba(100,0,255,0.12);
            text-align:center;
        ">
            <div style="
                color:#00eaff;
                font-size:13px;
                font-weight:800;
                letter-spacing:4px;
                margin-bottom:15px;
            ">
                ✦ JON AI DAILY MOTIVATION ✦
            </div>

            <div style="
                color:white;
                font-size:22px;
                font-weight:700;
                line-height:1.5;
                text-shadow:0 0 12px rgba(0,220,255,0.35);
            ">
                “{quote}”
            </div>

            <div style="
                color:#9ca9c7;
                font-size:12px;
                margin-top:17px;
                letter-spacing:2px;
            ">
                BUGUNGI MAQSADING — ERTANGI NATIJANG
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ------------------------------------------------------------
# 5. BOSH SAHIFA — MOTIVATSIYA RAMKASI BILAN YANGILANGAN
# Bu funksiya oldingi render_home() o‘rnini egallaydi.
# ------------------------------------------------------------

def render_home():

    st.markdown(
        """
        <div style="text-align:center; padding:10px 0;">
            <div style="
                font-size:38px;
                font-weight:900;
                letter-spacing:5px;
                color:#00eaff;
                text-shadow:0 0 18px rgba(0,220,255,0.65);
            ">
                JON AI
            </div>

            <div style="
                color:#9ba9c9;
                font-size:12px;
                letter-spacing:5px;
            ">
                LIVE INTELLIGENCE
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Asosiy yagona jonli personaj
    show_character(height=590)

    # Motivatsiya ramkasi
    render_motivation_frame()

    st.markdown("### ✦ JON AI imkoniyatlari")

    feature_columns = st.columns(2)

    features = [
        ("🤖", "AI Chat", "AI bilan suhbat"),
        ("🎓", "Ta'lim markazi", "Bilim va o‘rganish"),
        ("📊", "PowerPoint", "Taqdimot yaratish"),
        ("🌐", "Tarjimon", "Ko‘p tilli tarjima"),
        ("🎨", "Rasm prompti", "Ijodiy g‘oyalar"),
        ("📈", "Excel tahlil", "Ma'lumotlarni tahlil qilish")
    ]

    for index, feature in enumerate(features):

        icon, title, description = feature
        column = feature_columns[index % 2]

        with column:

            st.markdown(
                f"""
                <div style="
                    padding:17px;
                    margin:6px 0;
                    border-radius:15px;
                    background:rgba(15,30,60,0.75);
                    border:1px solid rgba(0,220,255,0.25);
                ">
                    <div style="font-size:25px;">{icon}</div>
                    <div style="
                        color:white;
                        font-weight:800;
                        margin-top:5px;
                    ">{title}</div>
                    <div style="
                        color:#9ba9c9;
                        font-size:12px;
                    ">{description}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

            if st.button(
                f"{title} →",
                key=f"home_feature_{index}",
                use_container_width=True
            ):
                go_to_page(title)
    
# ============================================================
# JON AI LIVE 10.0 — 5-QISM
# TA'LIM MARKAZI + POWERPOINT + TARJIMON
# ============================================================


# ------------------------------------------------------------
# 1. GEMINI BILAN ISHLASH UCHUN YORDAMCHI FUNKSIYA
# ------------------------------------------------------------

def jon_generate(prompt):

    try:
        result = ask_ai(prompt)

        if not result:
            return "JON hozir javob tayyorlay olmadi. Qayta urinib ko‘ring."

        return str(result).strip()

    except Exception as error:
        return explain_api_error(error)


# ------------------------------------------------------------
# 2. TA'LIM MARKAZI
# ------------------------------------------------------------

def render_education():

    st.title("🎓 JON AI — Ta'lim markazi")

    st.write(
        "Turli fanlarni o‘rganing, mavzularni tushuning "
        "va bilimlaringizni sinab ko‘ring."
    )

    tab1, tab2, tab3 = st.tabs([
        "📚 Mavzuni tushuntirish",
        "📝 Test yaratish",
        "✍️ Mashq va topshiriq"
    ])

    with tab1:

        st.subheader("Mavzuni sodda tushuntirish")

        subject = st.text_input(
            "Fan yoki mavzu",
            placeholder="Masalan: Ingliz tili — Present Perfect",
            key="edu_subject"
        )

        level = st.selectbox(
            "Bilim darajasi",
            [
                "Boshlang‘ich",
                "O‘rta",
                "Yuqori",
                "Professional"
            ],
            key="edu_level"
        )

        explanation_style = st.selectbox(
            "Tushuntirish uslubi",
            [
                "Sodda va tushunarli",
                "Batafsil va chuqur",
                "Misollar bilan",
                "Bosqichma-bosqich"
            ],
            key="edu_style"
        )

        if st.button(
            "Mavzuni tushuntir",
            key="edu_explain",
            use_container_width=True
        ):

            if not subject.strip():
                st.warning("Avval mavzuni kiriting.")
            else:

                prompt = f"""
                Quyidagi mavzuni o‘zbek tilida tushuntir.

                Mavzu: {subject}
                O‘quvchi darajasi: {level}
                Uslub: {explanation_style}

                Talablar:
                1. Mavzuning asosiy mazmunini tushuntir.
                2. Muhim qoidalarni ko‘rsat.
                3. Hayotiy misollar keltir.
                4. Keng tarqalgan xatolarni tushuntir.
                5. Qisqa xulosa ber.
                """

                with st.spinner("JON mavzuni tayyorlamoqda..."):
                    result = jon_generate(prompt)

                st.markdown("### 📖 Dars")
                st.markdown(result)

    with tab2:

        st.subheader("Aqlli test yaratuvchi")

        test_topic = st.text_input(
            "Test mavzusi",
            placeholder="Masalan: Ingliz tili grammatikasi",
            key="edu_test_topic"
        )

        question_count = st.slider(
            "Savollar soni",
            min_value=3,
            max_value=20,
            value=5,
            key="edu_question_count"
        )

        test_level = st.selectbox(
            "Test qiyinligi",
            ["Oson", "O‘rta", "Qiyin"],
            key="edu_test_level"
        )

        if st.button(
            "Test yaratish",
            key="edu_create_test",
            use_container_width=True
        ):

            if not test_topic.strip():
                st.warning("Test mavzusini kiriting.")
            else:

                prompt = f"""
                {test_topic} mavzusida {question_count} ta test yarat.

                Qiyinlik darajasi: {test_level}.

                Har bir savolda A, B, C, D variantlari bo‘lsin.
                Oxirida javoblar kalitini va qisqa izohlarni ber.
                Savollar bir-birini takrorlamasin.
                Javoblar ilmiy jihatdan to‘g‘ri bo‘lsin.
                """

                with st.spinner("Test yaratilmoqda..."):
                    result = jon_generate(prompt)

                st.markdown("### 📝 Tayyor test")
                st.markdown(result)

    with tab3:

        st.subheader("Mashq va topshiriqlar")

        exercise_topic = st.text_input(
            "Mashq mavzusi",
            placeholder="Masalan: Python dasturlash",
            key="edu_exercise_topic"
        )

        exercise_type = st.selectbox(
            "Topshiriq turi",
            [
                "Amaliy mashqlar",
                "Uyga vazifa",
                "Mustaqil ish",
                "Savol-javob",
                "Imtihonga tayyorgarlik"
            ],
            key="edu_exercise_type"
        )

        if st.button(
            "Topshiriq tayyorlash",
            key="edu_create_exercise",
            use_container_width=True
        ):

            if not exercise_topic.strip():
                st.warning("Mavzuni kiriting.")
            else:

                prompt = f"""
                {exercise_topic} mavzusida o‘zbek tilida
                quyidagi turdagi topshiriq tayyorla:

                Topshiriq turi: {exercise_type}

                Talablar:
                - Bosqichma-bosqich topshiriqlar ber.
                - Qiyinlik darajasini asta oshir.
                - Har bir topshiriq maqsadini tushuntir.
                - Oxirida tekshirish mezonlarini ber.
                """

                with st.spinner("Topshiriqlar tayyorlanmoqda..."):
                    result = jon_generate(prompt)

                st.markdown("### ✍️ Topshiriqlar")
                st.markdown(result)


# ------------------------------------------------------------
# 3. POWERPOINT FAYL YARATISH
# ------------------------------------------------------------

def create_powerpoint(title, slide_text):

    presentation = Presentation()

    # Birinchi slayd — sarlavha
    first_slide = presentation.slides.add_slide(
        presentation.slide_layouts[0]
    )

    first_slide.shapes.title.text = title

    if len(first_slide.placeholders) > 1:
        first_slide.placeholders[1].text = (
            "JON AI yordamida tayyorlandi"
        )

    # Matnni slaydlar bo‘yicha ajratish
    sections = re.split(
        r"(?im)^\s*(?:SLIDE|SLAYD)\s*\d*\s*[:.\-)]*\s*",
        slide_text
    )

    sections = [
        section.strip()
        for section in sections
        if section.strip()
    ]

    # Agar model SLIDE formatida javob bermasa,
    # matnni paragraflar bo‘yicha ajratamiz.
    if len(sections) <= 1:

        paragraphs = [
            paragraph.strip()
            for paragraph in slide_text.split("\n\n")
            if paragraph.strip()
        ]

        sections = paragraphs if paragraphs else [slide_text]

    for index, section in enumerate(sections[:20], start=1):

        lines = [
            line.strip()
            for line in section.splitlines()
            if line.strip()
        ]

        if not lines:
            continue

        slide_title = lines[0].strip("#* ")

        slide = presentation.slides.add_slide(
            presentation.slide_layouts[1]
        )

        slide.shapes.title.text = slide_title[:120]

        body = slide.placeholders[1].text_frame
        body.clear()

        content_lines = lines[1:] if len(lines) > 1 else lines

        for line_index, line in enumerate(content_lines):

            if line_index == 0:
                paragraph = body.paragraphs[0]
            else:
                paragraph = body.add_paragraph()

            paragraph.text = line.lstrip("-•* ").strip()
            paragraph.level = 0
            paragraph.font.size = Pt(20)

    output = BytesIO()
    presentation.save(output)
    output.seek(0)

    return output.getvalue()


# ------------------------------------------------------------
# 4. POWERPOINT SAHIFASI
# ------------------------------------------------------------

def render_powerpoint():

    st.title("📊 JON AI — PowerPoint Studio")

    st.write(
        "Mavzuni kiriting. JON AI siz uchun taqdimot "
        "matnini yaratadi va PowerPoint fayliga aylantiradi."
    )

    title = st.text_input(
        "Taqdimot sarlavhasi",
        placeholder="Masalan: Sun'iy intellekt kelajagi",
        key="ppt_title"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        min_value=3,
        max_value=20,
        value=7,
        key="ppt_slide_count"
    )

    audience = st.selectbox(
        "Taqdimot kimlar uchun?",
        [
            "Maktab o‘quvchilari",
            "Universitet talabalari",
            "O‘qituvchilar",
            "Biznes va kompaniyalar",
            "Umumiy auditoriya"
        ],
        key="ppt_audience"
    )

    design_style = st.selectbox(
        "Taqdimot uslubi",
        [
            "Professional",
            "Minimalistik",
            "Zamonaviy texnologik",
            "Ilmiy",
            "Biznes"
        ],
        key="ppt_style"
    )

    if st.button(
        "🚀 Taqdimot yaratish",
        key="ppt_generate",
        use_container_width=True
    ):

        if not title.strip():
            st.warning("Taqdimot mavzusini kiriting.")

        else:

            prompt = f"""
            "{title}" mavzusida {slide_count} ta slayd
            uchun taqdimot tayyorla.

            Auditoriya: {audience}
            Dizayn uslubi: {design_style}

            Natijani quyidagi formatda ber:

            SLIDE 1: Sarlavha
            Slayd matni

            SLIDE 2: Sarlavha
            Slayd matni

            Har bir slaydda:
            - Bitta aniq sarlavha.
            - 3-5 ta mazmunli fikr.
            - Zarur joyda misol yoki raqam.
            - Qisqa va tushunarli matn.

            Yakuniy slaydda xulosa bo‘lsin.
            O‘zbek tilida yoz.
            """

            with st.spinner("JON taqdimot ustida ishlamoqda..."):

                result = jon_generate(prompt)

            if result:

                st.session_state["ppt_result"] = result
                st.session_state["ppt_result_title"] = title

    if st.session_state.get("ppt_result"):

        st.markdown("### 📑 Taqdimot matni")

        st.markdown(st.session_state["ppt_result"])

        try:

            pptx_data = create_powerpoint(
                st.session_state.get(
                    "ppt_result_title",
                    "JON AI taqdimoti"
                ),
                st.session_state["ppt_result"]
            )

            st.download_button(
                label="⬇️ PowerPoint faylini yuklab olish",
                data=pptx_data,
                file_name="JON_AI_Presentation.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                key="ppt_download",
                use_container_width=True
            )

        except Exception as error:

            st.error(
                "PowerPoint faylini yaratishda xatolik: "
                + str(error)
            )


# ------------------------------------------------------------
# 5. KO‘P TILLI TARJIMON
# ------------------------------------------------------------

def render_translator():

    st.title("🌐 JON AI — Universal Translator")

    st.write(
        "Matnlarni turli tillarga tarjima qiling. "
        "JON AI tarjimaning mazmuni va ohangini saqlashga harakat qiladi."
    )

    languages = [
        "O‘zbek",
        "Ingliz",
        "Rus",
        "Turk",
        "Arab",
        "Nemis",
        "Fransuz",
        "Ispan",
        "Xitoy",
        "Yapon",
        "Koreys",
        "Italyan"
    ]

    col1, col2 = st.columns(2)

    with col1:
        source_language = st.selectbox(
            "Asl til",
            ["Avtomatik aniqlash"] + languages,
            key="translator_source"
        )

    with col2:
        target_language = st.selectbox(
            "Tarjima tili",
            languages,
            index=1,
            key="translator_target"
        )

    original_text = st.text_area(
        "Tarjima qilinadigan matn",
        height=200,
        placeholder="Matningizni shu yerga yozing...",
        key="translator_text"
    )

    translation_style = st.selectbox(
        "Tarjima uslubi",
        [
            "Tabiiy va ravon",
            "Rasmiy",
            "Akademik",
            "Oddiy kundalik",
            "So‘zma-so‘z"
        ],
        key="translator_style"
    )

    if st.button(
        "🌍 Tarjima qilish",
        key="translate_button",
        use_container_width=True
    ):

        if not original_text.strip():

            st.warning("Tarjima qilinadigan matnni kiriting.")

        else:

            prompt = f"""
            Quyidagi matnni tarjima qil.

            Asl til: {source_language}
            Maqsad til: {target_language}
            Uslub: {translation_style}

            Talablar:
            - Asl ma'noni saqla.
            - Grammatik jihatdan to‘g‘ri tarjima qil.
            - Keraksiz izohlar qo‘shma.
            - Ismlar va maxsus atamalarni ehtiyotkorlik bilan tarjima qil.
            - Agar so‘zning bir nechta ma'nosi bo‘lsa,
              kontekstga mosini tanla.

            MATN:
            {original_text}
            """

            with st.spinner("JON tarjima qilmoqda..."):

                translated = jon_generate(prompt)

            st.session_state["translation_result"] = translated

    if st.session_state.get("translation_result"):

        st.markdown("### ✅ Tarjima natijasi")

        st.text_area(
            "Tayyor tarjima",
            value=st.session_state["translation_result"],
            height=200,
            key="translation_output"
        )

        st.button(
            "📋 Tarjimani nusxalash",
            key="copy_translation",
            on_click=lambda: st.session_state.update({
                "translation_copied": True
            })
        )

        if st.session_state.get("translation_copied"):
            st.info(
                "Tarjima matni yuqoridagi maydonda. "
                "Uni bosib, matnni nusxalashingiz mumkin."
            )

# ============================================================
# 5-QISM TUGADI
# KEYINGI QISM: RASM PROMPTI + EXCEL TAHLIL + VAZIFALAR
# ============================================================
