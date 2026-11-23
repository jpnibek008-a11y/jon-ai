
# ============================================================
# JON AI LIVE 10.1 — YAXLIT, TARTIBLANGAN STREAMLIT ILOVASI
# ============================================================

import json
import random
import re
from datetime import date
from io import BytesIO

import streamlit as st
import streamlit.components.v1 as components

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from google import genai
except Exception:
    genai = None

try:
    from pptx import Presentation
    from pptx.util import Pt
except Exception:
    Presentation = None
    Pt = None

APP_NAME = "JON AI LIVE"
APP_VERSION = "10.1"

DEFAULT_MODEL = "gemini-3.5-flash-lite"
PAGES = [
    "🏠 Bosh sahifa", "🤖 AI Chat", "🎓 Ta'lim markazi",
    "📊 PowerPoint", "🌐 Tarjimon", "🎨 Rasm prompti",
    "📈 Excel tahlil", "📝 Vazifalar", "📷 Kamera",
    "👤 Personaj", "⚙️ Sozlamalar", "💎 Premium"
]

JON_CORE = """
Siz JON AI nomli professional yordamchisiz. O'zbek tilida aniq, foydali,
tushunarli javob bering. Bilmagan narsangizni to'qimang. Noaniqlikni ochiq
ayting. Maxfiy ma'lumotlarni so'ramang. Murakkab ishlarni bosqichlarga ajrating.
"""

MOTIVATION_QUOTES = [
    ("TO‘XTAMA. SEN HALI ORZU QILGAN MANZILINGGA YETMAGANSAN!", "Charchashing mumkin, ammo taslim bo‘lma."),
    ("BUGUNGI HARAKATING — ERTANGI G‘ALABANG.", "Kichik qadamlar katta natijalarga olib keladi."),
    ("QO‘RQMA. BOSHLAGAN ISHINGNI OXIRIGA YETKAZ!", "Mukammal bo‘lish shart emas. Muhimi — davom etish."),
    ("SEN O‘YLAGANINGDAN HAM KUCHLISAN!", "Qiyinchiliklar seni chiniqtiradi."),
    ("BUGUN TASLIM BO‘LMA!", "Ertangi o‘zing bugungi harakating uchun rahmat aytadi."),
]

DEFAULTS = {
    "messages": [], "tasks": [], "character_name": "JON",
    "character_style": "Hologramma", "character_animation": "Tabiiy harakat",
    "api_key": "", "model": DEFAULT_MODEL, "selected_page": PAGES[0],
    "jon_mood": "neutral", "jon_action": "speaking", "last_answer": "",
    "ppt_result": "", "ppt_result_title": "JON AI taqdimoti",
    "translation_result": "", "edu_result": "", "image_prompt_result": "",
    "excel_result": "", "motivation_index": date.today().toordinal() % len(MOTIVATION_QUOTES),
}

st.set_page_config(page_title=f"{APP_NAME} {APP_VERSION}", page_icon="💎", layout="wide")
for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value

# --------------------------- DESIGN ---------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Orbitron:wght@600;700;800&display=swap');
.stApp {background:radial-gradient(ellipse at 12% 8%,#17265a 0%,transparent 34%),radial-gradient(ellipse at 88% 18%,#25124d 0%,transparent 35%),linear-gradient(145deg,#050817,#080d20 60%,#09051c);color:#edf4ff;font-family:Inter,sans-serif}
h1,h2,h3 {color:#edf4ff!important}
.hero-title {font-family:Orbitron,sans-serif;text-align:center;font-size:clamp(30px,5vw,55px);font-weight:800;letter-spacing:4px;color:#8beaff;text-shadow:0 0 20px #26baff88}
.hero-subtitle {text-align:center;color:#9bb8df;letter-spacing:4px;font-size:11px}
.motivation-frame {padding:24px;margin:20px 0;text-align:center;border:1px solid #42cfff;border-radius:20px;background:linear-gradient(125deg,#091c3a,#201141);box-shadow:0 0 22px #26baff33}
.motivation-label {color:#65dcff;font-size:11px;font-weight:800;letter-spacing:3px}
.feature-card {padding:18px;border:1px solid #27466f;border-radius:16px;background:#0c1730;margin-bottom:10px}
.stButton>button {border-radius:12px;border:1px solid #36bdf2;background:linear-gradient(100deg,#123b68,#34205e);color:white;font-weight:700;min-height:42px}
.stTextInput input,.stTextArea textarea,.stSelectbox div[data-baseweb="select"] {background:#101c35;color:#edf4ff}
</style>
""", unsafe_allow_html=True)

# --------------------------- API ---------------------------

def get_api_key():
    key = st.session_state.get("api_key", "").strip()
    if key:
        return key
    try:
        return st.secrets.get("GEMINI_API_KEY", "").strip()
    except Exception:
        return ""

def gemini_text(prompt, system_instruction=None):
    if genai is None:
        return "Google GenAI kutubxonasi o‘rnatilmagan. requirements.txt faylini tekshiring."
    api_key = get_api_key()
    if not api_key:
        return "Gemini API kalitini Sozlamalar bo‘limiga kiriting."
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=st.session_state.get("model", DEFAULT_MODEL),
            contents=prompt,
            config={"system_instruction": system_instruction or JON_CORE}
        )
        return response.text or "AI javob qaytarmadi."
    except Exception as e:
        error = str(e)
        if "429" in error:
            return "Gemini limiti tugagan (429). Biroz kuting yoki billing va kvotani tekshiring."
        if "503" in error:
            return "Gemini serveri vaqtincha band (503). Birozdan keyin qayta urinib ko‘ring."
        if "401" in error or "403" in error:
            return "API kaliti noto‘g‘ri yoki ruxsati yo‘q. Sozlamalarni tekshiring."
        if "404" in error:
            return "Tanlangan Gemini modeli topilmadi. Sozlamalardan model nomini tekshiring."
        return f"Gemini bilan bog‘lanishda xatolik: {error}"

def ask_jon_with_behavior(prompt):
    behavior_prompt = f"""
Foydalanuvchi savoliga javob ber. Javobni quyidagi JSON formatida qaytar:
{{"answer":"javob matni","mood":"neutral|happy|thinking|serious|greeting","action":"speaking|thinking|idle"}}
Faqat JSON qaytar.
Savol: {prompt}
"""
    raw = gemini_text(behavior_prompt)
    try:
        match = re.search(r"\{.*\}", raw, re.S)
        data = json.loads(match.group(0)) if match else {}
        answer = data.get("answer", raw)
        mood = data.get("mood", "neutral")
        action = data.get("action", "speaking")
        if mood not in ["neutral", "happy", "thinking", "serious", "greeting"]:
            mood = "neutral"
        if action not in ["speaking", "thinking", "idle"]:
            action = "speaking"
        st.session_state["jon_mood"] = mood
        st.session_state["jon_action"] = action
        return answer
    except Exception:
        st.session_state["jon_mood"] = "neutral"
        st.session_state["jon_action"] = "speaking"
        return raw
        
# ---------------------- 3D JON CHARACTER ----------------------


def show_character(height=560):
    config = {
        "height": int(height),
        "name": st.session_state.get("character_name", "JON"),
        "style": st.session_state.get("character_style", "Hologramma"),
        "mode": st.session_state.get("character_animation", "Tabiiy harakat"),
        "mood": st.session_state.get("jon_mood", "neutral"),
        "action": st.session_state.get("jon_action", "speaking"),
    }

    config_json = json.dumps(
        config, ensure_ascii=False
    ).replace("</", "<\\/")

    html_code = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

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

#scene {
    width: 100%;
    height: CONFIG_HEIGHTpx;
    position: relative;
    overflow: hidden;
    border-radius: 22px;
    background:
      radial-gradient(ellipse at 50% 42%,
      #173c70 0%, #0a1936 48%, #030714 100%);
    border: 1px solid rgba(83,210,255,.25);
}

canvas {
    display: block;
    width: 100%;
    height: 100%;
}

#hud {
    position: absolute;
    top: 16px;
    left: 0;
    width: 100%;
    text-align: center;
    z-index: 5;
    pointer-events: none;
}

#name {
    color: #8ceeff;
    font-size: 25px;
    font-weight: 900;
    letter-spacing: 7px;
    text-shadow:
      0 0 8px #00cfff,
      0 0 25px #008cff;
}

#status {
    color: #75ffd6;
    font-size: 10px;
    letter-spacing: 3px;
    margin-top: 7px;
}

#bottom {
    position: absolute;
    bottom: 15px;
    left: 0;
    width: 100%;
    text-align: center;
    color: #9edaff;
    font-size: 10px;
    letter-spacing: 2px;
    z-index: 5;
    pointer-events: none;
}

#glow {
    position: absolute;
    width: 240px;
    height: 240px;
    left: calc(50% - 120px);
    top: 35%;
    border-radius: 50%;
    background: #00aaff;
    filter: blur(95px);
    opacity: .18;
    pointer-events: none;
}

#scanlines {
    position: absolute;
    inset: 0;
    z-index: 3;
    pointer-events: none;
    opacity: .12;
    background: repeating-linear-gradient(
        to bottom,
        transparent 0px,
        transparent 3px,
        rgba(95,220,255,.2) 4px
    );
}

#corner {
    position: absolute;
    inset: 8px;
    border: 1px solid rgba(75,210,255,.12);
    border-radius: 17px;
    pointer-events: none;
    z-index: 4;
}

#fallback {
    display: none;
    position: absolute;
    inset: 0;
    color: #8ceeff;
    align-items: center;
    justify-content: center;
    text-align: center;
    font-size: 14px;
    z-index: 10;
}
</style>
</head>

<body>

<div id="scene">
    <div id="glow"></div>
    <div id="scanlines"></div>
    <div id="corner"></div>

    <div id="hud">
        <div id="name"></div>
        <div id="status">● JON AI — ONLINE</div>
    </div>

    <div id="bottom"></div>
    <div id="fallback">3D modul yuklanmoqda...</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>

<script>
(function(){

    const CFG = CONFIG_JSON;

    const root = document.getElementById("scene");
    const status = document.getElementById("status");
    const fallback = document.getElementById("fallback");

    document.getElementById("name").textContent = CFG.name;
    document.getElementById("bottom").textContent =
        CFG.style + " · " + CFG.mode;

    if (!window.THREE) {
        fallback.style.display = "flex";
        fallback.textContent = "3D modul yuklanmadi. Internet aloqasini tekshiring.";
        status.textContent = "3D OFFLINE";
        return;
    }

    const THREE = window.THREE;

    const scene = new THREE.Scene();

    const camera = new THREE.PerspectiveCamera(
        38,
        root.clientWidth / root.clientHeight,
        0.1,
        100
    );

    camera.position.set(0, 0.3, 8.5);
    camera.lookAt(0, 0, 0);

    const renderer = new THREE.WebGLRenderer({
        antialias: true,
        alpha: true,
        powerPreference: "high-performance"
    });

    renderer.setPixelRatio(
        Math.min(window.devicePixelRatio || 1, 1.6)
    );

    renderer.setSize(root.clientWidth, root.clientHeight);
    renderer.setClearColor(0x000000, 0);

    renderer.outputEncoding = THREE.sRGBEncoding;

    root.insertBefore(renderer.domElement, root.firstChild);

    const character = new THREE.Group();
    scene.add(character);

    // ======================================
    // STYLE SYSTEM
    // ======================================

    const style = CFG.style;

    const hologram = style === "Hologramma";
    const human = style === "Human" ||
                  style === "Realistik inson";
    const android = style === "Android" ||
                    style === "Realistik robot";

    const cartoon = style === "Multfilm uslubi";
    const fantasy = style === "Fantastik";

    const primaryColor = fantasy ? 0xb15cff :
                         hologram ? 0x25dfff :
                         cartoon ? 0x36dfff :
                         0x319cff;

    const skinColor = human ? 0xdcae91 : 0xb8d4e8;

    const bodyColor = human ? 0x172b48 :
                      hologram ? 0x26cfff :
                      fantasy ? 0x6939a8 :
                      0x52718f;

    function makeMaterial(color, metalness, roughness, emissive) {
        return new THREE.MeshStandardMaterial({
            color: color,
            metalness: metalness,
            roughness: roughness,
            emissive: emissive || 0x000000,
            emissiveIntensity: hologram ? 0.75 : 0.25,
            transparent: hologram,
            opacity: hologram ? 0.76 : 1
        });
    }

    const bodyMat = makeMaterial(
        bodyColor, human ? 0.15 : 0.8, 0.3, primaryColor
    );

    const darkMat = makeMaterial(0x101a30, 0.7, 0.3);
    const skinMat = makeMaterial(skinColor, 0.05, 0.75);
    const metalMat = makeMaterial(0x91b9d5, 0.85, 0.22);

    const lightMat = new THREE.MeshBasicMaterial({
        color: primaryColor
    });

    const eyeMat = new THREE.MeshBasicMaterial({
        color: 0x8ef7ff
    });

    const whiteEyeMat = new THREE.MeshStandardMaterial({
        color: 0xf2f7ff,
        roughness: 0.3
    });

    // ======================================
    // LIGHTING
    // ======================================

    scene.add(new THREE.AmbientLight(0x9bcaff, 1.5));

    const blueLight = new THREE.PointLight(
        0x00bfff, 3.5, 18
    );
    blueLight.position.set(-3, 3, 4);
    scene.add(blueLight);

    const purpleLight = new THREE.PointLight(
        0x9b48ff, 2.5, 18
    );
    purpleLight.position.set(3, 1, 3);
    scene.add(purpleLight);

    const frontLight = new THREE.PointLight(
        0xffffff, 1.5, 12
    );
    frontLight.position.set(0, 3, 5);
    scene.add(frontLight);

    // ======================================
    // GEOMETRY HELPERS
    // ======================================

    function sphere(parent, x, y, z, sx, sy, sz, mat) {
        const mesh = new THREE.Mesh(
            new THREE.SphereGeometry(1, 32, 24),
            mat
        );

        mesh.position.set(x, y, z);
        mesh.scale.set(sx, sy, sz);
        parent.add(mesh);
        return mesh;
    }

    function cylinderBetween(parent, start, end, radius, mat) {
        const direction = new THREE.Vector3().subVectors(end, start);
        const length = direction.length();

        const mesh = new THREE.Mesh(
            new THREE.CylinderGeometry(
                radius * 0.85,
                radius,
                length,
                20
            ),
            mat
        );

        mesh.position.copy(start).add(end).multiplyScalar(0.5);
        mesh.quaternion.setFromUnitVectors(
            new THREE.Vector3(0, 1, 0),
            direction.normalize()
        );

        parent.add(mesh);
        return mesh;
    }

    // ======================================
    // TORSO
    // ======================================

    const torso = sphere(
        character, 0, -0.35, 0,
        0.67, 0.88, 0.37, bodyMat
    );

    const chest = sphere(
        character, 0, -0.20, 0.345,
        0.35, 0.31, 0.055, lightMat
    );

    sphere(
        character, -0.69, 0.16, 0,
        0.28, 0.26, 0.28, darkMat
    );

    sphere(
        character, 0.69, 0.16, 0,
        0.28, 0.26, 0.28, darkMat
    );

    // CHEST CORE
    const core = sphere(
        character, 0, -0.35, 0.42,
        0.12, 0.12, 0.045, eyeMat
    );

    // NECK
    cylinderBetween(
        character,
        new THREE.Vector3(0, 0.38, 0),
        new THREE.Vector3(0, 0.83, 0),
        0.17,
        darkMat
    );

    // ======================================
    // HEAD AND FACE
    // ======================================

    const headGroup = new THREE.Group();
    headGroup.position.set(0, 1.25, 0);
    character.add(headGroup);

    const headMat = human ? skinMat : bodyMat;

    const head = sphere(
        headGroup, 0, 0, 0,
        human ? 0.48 : 0.53,
        0.64,
        0.43,
        headMat
    );

    // Hair for human character
    if (human) {
        const hairMat = makeMaterial(0x25202a, 0.05, 0.85);

        sphere(
            headGroup, 0, 0.38, -0.035,
            0.49, 0.30, 0.44, hairMat
        );

        sphere(
            headGroup, 0, 0.21, 0.29,
            0.39, 0.12, 0.13, hairMat
        );
    }

    // Face panel for android
    if (android || hologram) {
        sphere(
            headGroup, 0, -0.045, 0.37,
            0.40, 0.39, 0.11, darkMat
        );
    }

    // Ears
    for (const side of [-1, 1]) {
        sphere(
            headGroup, side * 0.48, -0.04, 0,
            0.105, 0.19, 0.13,
            android ? metalMat : headMat
        );

        if (android || hologram) {
            sphere(
                headGroup, side * 0.52, -0.04, 0.10,
                0.045, 0.07, 0.045, lightMat
            );
        }
    }

    // Eyes
    const eyes = [];

    for (const side of [-1, 1]) {

        const eyeGroup = new THREE.Group();
        eyeGroup.position.set(side * 0.19, 0.045, 0.405);
        headGroup.add(eyeGroup);

        sphere(
            eyeGroup, 0, 0, 0,
            0.105, 0.062, 0.035,
            human ? whiteEyeMat : eyeMat
        );

        sphere(
            eyeGroup, 0, 0, 0.028,
            0.038, 0.043, 0.018,
            human ? darkMat : lightMat
        );

        sphere(
            eyeGroup, 0.012, 0.014, 0.045,
            0.012, 0.012, 0.009,
            whiteEyeMat
        );

        eyes.push(eyeGroup);

        // Eyebrows
        if (human) {
            const brow = sphere(
                headGroup,
                side * 0.19, 0.145, 0.415,
                0.105, 0.022, 0.025,
                darkMat
            );
            brow.rotation.z = side * -0.08;
        }
    }

    // Nose
    if (human) {
        sphere(
            headGroup, 0, -0.085, 0.42,
            0.065, 0.10, 0.075, skinMat
        );
    }

    // Mouth
    const mouth = sphere(
        headGroup, 0, -0.245, 0.405,
        0.11, 0.018, 0.025,
        human ? darkMat : lightMat
    );

    // Smile line
    const smile = new THREE.Mesh(
        new THREE.TorusGeometry(0.09, 0.009, 6, 24, Math.PI),
        new THREE.MeshBasicMaterial({
            color: human ? 0x8c4b4b : primaryColor
        })
    );

    smile.position.set(0, -0.23, 0.42);
    smile.rotation.z = Math.PI;
    headGroup.add(smile);

    // ======================================
    // ARMS
    // ======================================

    const arms = [];

    for (const side of [-1, 1]) {

        const pivot = new THREE.Group();
        pivot.position.set(side * 0.68, 0.22, 0);
        character.add(pivot);

        const upper = sphere(
            pivot, side * 0.02, -0.34, 0,
            0.16, 0.42, 0.17,
            bodyMat
        );

        const elbow = sphere(
            pivot, side * 0.03, -0.68, 0,
            0.14, 0.14, 0.14,
            metalMat
        );

        sphere(
            pivot, side * 0.04, -0.91, 0.01,
            0.14, 0.34, 0.14,
            bodyMat
        );

        sphere(
            pivot, side * 0.04, -1.18, 0.05,
            0.16, 0.17, 0.16,
            human ? skinMat : metalMat
        );

        if (android || hologram) {
            sphere(
                pivot, side * 0.04, -0.91, 0.14,
                0.06, 0.13, 0.025, lightMat
            );
        }

        arms.push(pivot);
    }

    // ======================================
    // LEGS
    // ======================================

    const legs = [];

    for (const side of [-1, 1]) {

        const pivot = new THREE.Group();
        pivot.position.set(side * 0.30, -0.93, 0);
        character.add(pivot);

        sphere(
            pivot, 0, -0.35, 0,
            0.24, 0.43, 0.25,
            bodyMat
        );

        sphere(
            pivot, 0, -0.79, 0,
            0.19, 0.43, 0.19,
            metalMat
        );

        sphere(
            pivot, 0, -1.10, 0.12,
            0.22, 0.13, 0.34,
            darkMat
        );

        legs.push(pivot);
    }

    // ======================================
    // ENERGY RINGS
    // ======================================

    const rings = [];

    for (let i = 0; i < 3; i++) {

        const ring = new THREE.Mesh(
            new THREE.TorusGeometry(
                1.15 + i * 0.20,
                0.012,
                8,
                100
            ),
            new THREE.MeshBasicMaterial({
                color: i === 1 ? 0xb15cff : primaryColor,
                transparent: true,
                opacity: 0.7 - i * 0.12
            })
        );

        ring.rotation.x = Math.PI / 2 + i * 0.22;
        ring.position.y = -2.08 + i * 0.12;

        character.add(ring);
        rings.push(ring);
    }

    // ======================================
    // PARTICLES
    // ======================================

    const particleCount = 130;
    const positions = new Float32Array(particleCount * 3);

    for (let i = 0; i < particleCount; i++) {
        positions[i * 3] = (Math.random() - 0.5) * 5;
        positions[i * 3 + 1] = (Math.random() - 0.5) * 5;
        positions[i * 3 + 2] = (Math.random() - 0.5) * 3;
    }

    const particleGeometry = new THREE.BufferGeometry();

    particleGeometry.setAttribute(
        "position",
        new THREE.BufferAttribute(positions, 3)
    );

    const particles = new THREE.Points(
        particleGeometry,
        new THREE.PointsMaterial({
            color: primaryColor,
            size: 0.035,
            transparent: true,
            opacity: 0.85
        })
    );

    scene.add(particles);

    // ======================================
    // BASE
    // ======================================

    const base = new THREE.Mesh(
        new THREE.CylinderGeometry(1.05, 1.25, 0.10, 48),
        new THREE.MeshBasicMaterial({
            color: 0x164c85,
            transparent: true,
            opacity: 0.65
        })
    );

    base.position.y = -2.18;
    character.add(base);

    // ======================================
    // ANIMATION ENGINE
    // ======================================

    const clock = new THREE.Clock();

    let animationFrame = null;
    let active = true;

    function animate() {

        if (!active) return;

        animationFrame = requestAnimationFrame(animate);

        const t = clock.getElapsedTime();

        const mode = CFG.mode;
        const mood = CFG.mood;
        const action = CFG.action;

        // Natural floating
        character.position.y = Math.sin(t * 1.5) * 0.045;

        // BODY ANIMATION
        if (mode === "Raqs") {

            character.rotation.y = Math.sin(t * 2) * 0.28;
            character.rotation.z = Math.sin(t * 3) * 0.07;

            arms[0].rotation.z = -0.45 + Math.sin(t * 4) * 0.5;
            arms[1].rotation.z = 0.45 + Math.cos(t * 4) * 0.5;

            legs[0].rotation.x = Math.sin(t * 4) * 0.20;
            legs[1].rotation.x = Math.cos(t * 4) * 0.20;

        } else if (mode === "Energetik") {

            character.rotation.y = Math.sin(t * 2.5) * 0.18;

            arms[0].rotation.z = -0.4 + Math.sin(t * 3) * 0.3;
            arms[1].rotation.z = 0.4 + Math.cos(t * 3) * 0.3;

        } else if (mode === "Sokin") {

            character.rotation.y = Math.sin(t * 0.5) * 0.06;
            arms[0].rotation.z = -0.12;
            arms[1].rotation.z = 0.12;

        } else {

            character.rotation.y = Math.sin(t * 0.8) * 0.12;

            arms[0].rotation.z = -0.15 + Math.sin(t * 1.3) * 0.08;
            arms[1].rotation.z = 0.15 + Math.cos(t * 1.3) * 0.08;
        }

        // GREETING
        if (action === "greeting" || mood === "greeting") {
            arms[1].rotation.z = 0.65 + Math.sin(t * 6) * 0.30;
        }

        // HEAD MOVEMENT
        headGroup.rotation.y = Math.sin(t * 0.8) * 0.12;
        headGroup.rotation.x = Math.sin(t * 0.7) * 0.045;

        if (mood === "thinking") {
            headGroup.rotation.x = -0.15 + Math.sin(t) * 0.035;
        }

        if (mood === "happy" || mood === "greeting") {
            headGroup.rotation.z = Math.sin(t * 2) * 0.035;
        } else {
            headGroup.rotation.z = Math.sin(t * 0.7) * 0.015;
        }

        // BLINKING
        const blink = Math.sin(t * 1.7) > 0.985 ? 0.12 : 1;

        eyes.forEach(eye => {
            eye.scale.y = blink;
        });

        // SPEAKING
        const speaking = action === "speaking";

        if (speaking) {
            mouth.scale.y = 1 + Math.abs(Math.sin(t * 11)) * 4;
            mouth.scale.x = 1 + Math.abs(Math.sin(t * 8)) * 0.5;
        } else {
            mouth.scale.y = 1;
            mouth.scale.x = 1;
        }

        // SMILE
        smile.visible = mood === "happy" || mood === "greeting";

        // ENERGY RINGS
        rings.forEach((ring, i) => {
            ring.rotation.z =
                t * (0.25 + i * 0.12) *
                (i % 2 === 0 ? 1 : -1);
        });

        // PARTICLES
        particles.rotation.y = t * 0.025;
        particles.rotation.x = Math.sin(t * 0.3) * 0.08;

        // LIGHT PULSE
        blueLight.intensity = 3 + Math.sin(t * 2) * 0.7;
        core.scale.setScalar(1 + Math.sin(t * 3) * 0.12);

        renderer.render(scene, camera);
    }

    animate();

    // ======================================
    // RESPONSIVE SCREEN
    // ======================================

    function resizeScene() {

        const width = root.clientWidth;
        const height = root.clientHeight;

        if (!width || !height) return;

        camera.aspect = width / height;
        camera.updateProjectionMatrix();

        renderer.setSize(width, height);
    }

    window.addEventListener("resize", resizeScene);

    if (window.ResizeObserver) {
        const observer = new ResizeObserver(resizeScene);
        observer.observe(root);
    }

    // ======================================
    // CLEANUP
    // ======================================

    window.addEventListener("pagehide", function() {
        active = false;

        if (animationFrame !== null) {
            cancelAnimationFrame(animationFrame);
        }

        renderer.dispose();
    });

})();
</script>

</div>
</body>
</html>
"""

    html_code = html_code.replace(
        "CONFIG_HEIGHT", str(int(height))
    )

    html_code = html_code.replace(
        "CONFIG_JSON", config_json
    )

    components.html(
        html_code,
        height=int(height) + 10,
        scrolling=False
    )
    
# ---------------------- NAVIGATION ----------------------------

def go_to_page(page):
    st.session_state["selected_page"] = page
    st.rerun()


def render_navigation():
    st.sidebar.markdown(
        "<div style='text-align:center;font-size:28px;font-weight:800;color:#73ddff'>💎 JON AI</div>",
        unsafe_allow_html=True
    )
    st.sidebar.caption("LIVE INTELLIGENCE")
    st.sidebar.divider()

    current = st.session_state.get("selected_page", PAGES[0])
    if current not in PAGES:
        current = PAGES[0]

    selected = st.sidebar.radio(
        "NAVIGATSIYA",
        PAGES,
        index=PAGES.index(current),
        key="jon_navigation"
    )

    st.session_state["selected_page"] = selected

    st.sidebar.divider()
    st.sidebar.markdown("**JON AI SYSTEM**")
    st.sidebar.caption(f"Versiya: {APP_VERSION}")
    st.sidebar.caption("Holat: 🟢 Ishga tayyor")

    if st.sidebar.button("🔄 Bosh sahifaga qaytish", use_container_width=True):
        go_to_page(PAGES[0])


# ---------------------- HOME PAGE ----------------------

def render_home():
    st.markdown(
        "<div class='hero-title'>JON AI</div>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<div class='hero-subtitle'>LIVE INTELLIGENCE · YOUR DIGITAL PARTNER</div>",
        unsafe_allow_html=True
    )

    st.markdown("")
    show_character(height=470)

    quote_title, quote_text = MOTIVATION_QUOTES[
        st.session_state["motivation_index"]
    ]

    st.markdown(
        f"""
        <div class="motivation-frame">
            <div class="motivation-label">✦ BUGUNGI MOTIVATSIYA ✦</div>
            <h2 style="color:#8ceaff;margin-top:16px">{quote_title}</h2>
            <p style="color:#d2e5ff;font-size:15px;margin-top:12px">
                {quote_text}
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("✨ Yangi motivatsiya", use_container_width=True):
        current = st.session_state["motivation_index"]
        new_index = (current + 1) % len(MOTIVATION_QUOTES)
        st.session_state["motivation_index"] = new_index
        st.rerun()

    st.markdown("## 🚀 JON AI imkoniyatlari")
    st.caption("Kerakli bo‘limni tanlang va ishni boshlang.")

    feature_pages = [
        ("🤖 AI Chat", "Sun’iy intellekt bilan suhbat"),
        ("🎓 Ta'lim markazi", "Dars, test va mashqlar"),
        ("📊 PowerPoint", "Taqdimot tayyorlash"),
        ("🌐 Tarjimon", "Zamonaviy va qadimgi tillar"),
        ("🎨 Rasm prompti", "Rasm yaratish uchun prompt"),
        ("📈 Excel tahlil", "Jadval va ma’lumotlarni tahlil"),
        ("📝 Vazifalar", "Kundalik vazifalarni boshqarish"),
        ("📷 Kamera", "Kamera orqali jonli tasvir"),
        ("👤 Personaj", "JON personajini sozlash"),
        ("💎 Premium", "Premium imkoniyatlar")
    ]

    for start in range(0, len(feature_pages), 2):
        col1, col2 = st.columns(2)
        pair = feature_pages[start:start + 2]

        with col1:
            page, description = pair[0]
            st.markdown(
                f"""
                <div class="feature-card">
                    <h3>{page}</h3>
                    <p style="color:#9bb8df">{description}</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button(
                f"{page} bo‘limini ochish",
                key=f"home_{start}",
                use_container_width=True
            ):
                go_to_page(page)

        if len(pair) > 1:
            with col2:
                page, description = pair[1]
                st.markdown(
                    f"""
                    <div class="feature-card">
                        <h3>{page}</h3>
                        <p style="color:#9bb8df">{description}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                if st.button(
                    f"{page} bo‘limini ochish",
                    key=f"home_{start+1}",
                    use_container_width=True
                ):
                    go_to_page(page)

    st.markdown("---")
    st.markdown(
        "<div style='text-align:center;color:#7896b8;font-size:12px'>"
        "JON AI LIVE · Kelajak texnologiyalari bir platformada"
        "</div>",
        unsafe_allow_html=True
            )
    
# ---------------------- AI CHAT ----------------------

def render_chat():
    st.title("🤖 JON AI — AI Chat")
    st.caption("Sizning aqlli raqamli yordamchingiz")

    col1, col2 = st.columns([3, 1])
    with col1:
        st.info("Savolingizni yozing. JON AI sizga yordam beradi.")
    with col2:
        if st.button("🗑️ Suhbatni tozalash", use_container_width=True):
            st.session_state["messages"] = []
            st.rerun()

    # Oldingi xabarlarni chiqarish
    for message in st.session_state["messages"]:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    prompt = st.chat_input("JON AI ga xabar yozing...")

    if prompt:
        st.session_state["messages"].append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI javob tayyorlamoqda..."):
                answer = ask_jon_with_behavior(prompt)
            st.markdown(answer)

        st.session_state["messages"].append({
            "role": "assistant",
            "content": answer
        })


# ---------------------- EDUCATION CENTER ----------------------

def render_education():
    st.title("🎓 JON AI — Ta'lim markazi")
    st.caption("O‘rganish, mashq qilish va bilimni sinash markazi")

    mode = st.selectbox(
        "Nima tayyorlaymiz?",
        [
            "📚 Mavzuni tushuntirish",
            "📝 Test savollari",
            "✍️ Amaliy mashqlar",
            "📖 Qisqa konspekt",
            "🎯 Imtihonga tayyorgarlik"
        ],
        key="education_mode"
    )

    subject = st.text_input(
        "Fan yoki yo‘nalish",
        placeholder="Masalan: Ingliz tili, matematika, dasturlash..."
    )

    topic = st.text_input(
        "Mavzu",
        placeholder="Masalan: Present Perfect, Algebra, Python..."
    )

    level = st.selectbox(
        "Bilim darajasi",
        ["Boshlang‘ich", "O‘rta", "Yuqori", "Mutaxassis"],
        key="education_level"
    )

    count = st.slider(
        "Savol yoki mashqlar soni",
        min_value=3,
        max_value=20,
        value=5,
        key="education_count"
    )

    extra = st.text_area(
        "Qo‘shimcha talab (ixtiyoriy)",
        placeholder="Masalan: Javoblarni oxirida alohida ko‘rsat..."
    )

    if st.button("🚀 Darsni tayyorlash", use_container_width=True):
        if not subject.strip() or not topic.strip():
            st.warning("Fan va mavzu nomini kiriting.")
        else:
            if mode == "📚 Mavzuni tushuntirish":
                task = (
                    "Mavzuni sodda va tushunarli qilib tushuntir. "
                    "Ta'rif, qoidalar, misollar va xulosa ber."
                )
            elif mode == "📝 Test savollari":
                task = (
                    f"{count} ta test savoli tuz. Har bir savolda "
                    "4 ta variant bo‘lsin. To‘g‘ri javoblarni oxirida "
                    "alohida javoblar kaliti sifatida ber."
                )
            elif mode == "✍️ Amaliy mashqlar":
                task = (
                    f"{count} ta amaliy mashq tuz. Osondan qiyinga "
                    "qarab tartibla. Oxirida javoblar yoki yechimlarni ber."
                )
            elif mode == "📖 Qisqa konspekt":
                task = (
                    "Mavzu bo‘yicha tartibli, muhim tushunchalar, "
                    "asosiy qoidalar va misollardan iborat konspekt yoz."
                )
            else:
                task = (
                    f"Imtihonga tayyorgarlik uchun {count} ta savol, "
                    "muhim qoidalar va amaliy tavsiyalar tayyorla."
                )

            prompt = f"""
Fan: {subject}
Mavzu: {topic}
Daraja: {level}
Vazifa: {task}
Qo‘shimcha talab: {extra}

Javobni o‘zbek tilida, tartibli va tushunarli yoz.
"""

            with st.spinner("Ta'lim materiali tayyorlanmoqda..."):
                result = gemini_text(prompt)

            st.session_state["edu_result"] = result

    if st.session_state.get("edu_result"):
        st.markdown("---")
        st.subheader("📘 Tayyorlangan material")
        st.markdown(st.session_state["edu_result"])

        st.download_button(
            "⬇️ Materialni TXT formatida yuklab olish",
            data=st.session_state["edu_result"],
            file_name="jon_ai_talim_materiali.txt",
            mime="text/plain",
            use_container_width=True
    )
    
# ---------------------- POWERPOINT GENERATOR ----------------------

def render_powerpoint():
    st.title("📊 JON AI — PowerPoint Studio")
    st.caption("AI yordamida taqdimot rejasini va PPTX faylini yaratish")

    title = st.text_input(
        "Taqdimot mavzusi",
        placeholder="Masalan: Sun’iy intellektning kelajagi"
    )

    count = st.slider(
        "Slaydlar soni",
        min_value=3,
        max_value=20,
        value=8,
        key="ppt_count"
    )

    audience = st.selectbox(
        "Kimlar uchun?",
        ["Maktab o‘quvchilari", "Talabalar", "O‘qituvchilar", "Mutaxassislar", "Rahbariyat"]
    )

    style = st.selectbox(
        "Taqdimot uslubi",
        ["Professional", "Ilmiy", "Minimalistik", "Zamonaviy", "Ijodiy"]
    )

    if st.button("✨ Taqdimot yaratish", use_container_width=True):
        if not title.strip():
            st.warning("Taqdimot mavzusini kiriting.")
        else:
            prompt = f"""
            "{title}" mavzusida {count} ta slayd uchun taqdimot tayyorla.

            Auditoriya: {audience}
            Uslub: {style}

            Har bir slayd uchun:
            1. Slayd sarlavhasi
            2. 3–5 ta muhim fikr
            3. Zarur bo‘lsa, misol yoki statistik ma'lumot
            4. Spiker uchun qisqa izoh

            Kirish, asosiy qism va xulosa bo‘lsin.
            Uydirma statistik ma'lumot kiritma.
            Javobni SLAYD 1, SLAYD 2 shaklida tartibla.
            """

            with st.spinner("Taqdimot rejasi tayyorlanmoqda..."):
                result = gemini_text(prompt)

            st.session_state["ppt_result"] = result
            st.session_state["ppt_result_title"] = title

    result = st.session_state.get("ppt_result", "")

    if result:
        st.markdown("---")
        st.subheader("📑 Taqdimot rejasi")
        st.markdown(result)

        st.download_button(
            "⬇️ Rejani TXT formatida yuklab olish",
            data=result,
            file_name="jon_ai_taqdimot_rejasi.txt",
            mime="text/plain",
            use_container_width=True
        )

        if Presentation is None:
            st.warning(
                "PPTX yaratish uchun requirements.txt faylida "
                "python-pptx kutubxonasi bo‘lishi kerak."
            )
        else:
            if st.button("📥 PPTX faylini yaratish", use_container_width=True):
                try:
                    prs = Presentation()

                    # Birinchi slayd
                    slide = prs.slides.add_slide(prs.slide_layouts[0])
                    slide.shapes.title.text = st.session_state[
                        "ppt_result_title"
                    ]
                    slide.placeholders[1].text = "JON AI yordamida tayyorlandi"

                    # AI rejasini slaydlarga ajratish
                    sections = re.split(
                        r"(?=SLAYD\s*\d+)",
                        result,
                        flags=re.IGNORECASE
                    )

                    for section in sections:
                        section = section.strip()
                        if not section:
                            continue

                        lines = [
                            line.strip()
                            for line in section.splitlines()
                            if line.strip()
                        ]

                        if not lines:
                            continue

                        slide = prs.slides.add_slide(prs.slide_layouts[1])

                        heading = re.sub(
                            r"^SLAYD\s*\d+\s*[:\-–]?\s*",
                            "",
                            lines[0],
                            flags=re.IGNORECASE
                        )

                        slide.shapes.title.text = heading[:150]

                        body = slide.placeholders[1].text_frame
                        body.clear()

                        for line in lines[1:]:
                            clean_line = re.sub(r"^[•\-*]\s*", "", line)
                            paragraph = body.add_paragraph()
                            paragraph.text = clean_line
                            paragraph.level = 0
                            paragraph.font.size = Pt(18)

                    output = BytesIO()
                    prs.save(output)
                    output.seek(0)

                    st.download_button(
                        "⬇️ Tayyor PowerPointni yuklab olish",
                        data=output.getvalue(),
                        file_name="JON_AI_Taqdimot.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"PPTX yaratishda xatolik: {e}")


# ---------------------- TRANSLATOR ----------------------

def render_translator():
    st.title("🌐 JON AI — Universal Tarjimon")
    st.caption(
        "Zamonaviy, tarixiy va qadimgi tillar uchun tarjima yordamchisi"
    )

    modern_languages = [
        "O‘zbek", "Ingliz", "Rus", "Turk", "Arab", "Nemis",
        "Fransuz", "Ispan", "Italyan", "Xitoy", "Yapon",
        "Koreys", "Fors", "Hind", "Portugal", "Ukrain"
    ]

    historical_languages = [
        "Lotin tili",
        "Qadimgi yunon tili",
        "Sanskrit",
        "Qadimgi turkiy (O‘rxun-Enasoy)",
        "Chig‘atoy / eski o‘zbek tili",
        "Qadimgi fors tili",
        "Avesto tili",
        "Qadimgi Misr tili",
        "Akkad tili",
        "Shumer tili",
        "Qadimgi ibroniy tili",
        "Suriyalik (klassik suryoniy) tili",
        "Got tili",
        "Qadimgi ingliz tili",
        "Qadimgi cherkov slavyan tili"
    ]

    language_group = st.radio(
        "Til turini tanlang",
        ["Zamonaviy tillar", "Qadimgi va tarixiy tillar"],
        horizontal=True
    )

    available_languages = (
        modern_languages
        if language_group == "Zamonaviy tillar"
        else historical_languages
    )

    col1, col2 = st.columns(2)

    with col1:
        source_language = st.selectbox(
            "Qaysi tildan?",
            ["Avtomatik aniqlash"] + available_languages,
            key="translator_source"
        )

    with col2:
        target_language = st.selectbox(
            "Qaysi tilga?",
            available_languages,
            index=0,
            key="translator_target"
        )

    translation_mode = st.selectbox(
        "Tarjima usuli",
        [
            "Tabiiy va ravon tarjima",
            "So‘zma-so‘z tarjima",
            "Grammatik tahlil",
            "Transliteratsiya",
            "Tarixiy-madaniy izoh"
        ]
    )

    source_text = st.text_area(
        "Tarjima qilinadigan matn",
        height=180,
        placeholder="Matn yoki so‘zni shu yerga yozing..."
    )

    if st.button("🌍 Tarjima qilish", use_container_width=True):
        if not source_text.strip():
            st.warning("Tarjima qilinadigan matnni kiriting.")
        else:
            historical_warning = ""

            if language_group == "Qadimgi va tarixiy tillar":
                historical_warning = """
                MUHIM TALABLAR:
                - Qadimgi tillar lug‘ati va manbalari cheklangan bo‘lishi mumkin.
                - Ishonchingiz komil bo‘lmagan so‘z yoki grammatik shaklni
                  to‘qib chiqarmang.
                - Noaniq joylarni [noaniq] deb belgilang.
                - Bir nechta ehtimoliy talqin bo‘lsa, ularni alohida ko‘rsating.
                - Tarixiy imlo, davr va dialekt farqlarini hisobga oling.
                - Matnni zamonaviy tilga moslab yubormang.
                - Dalil va taxminni bir-biridan ajrating.
                """

            prompt = f"""
            Siz professional tarjimon va tilshunos yordamchisisiz.

            Manba tili: {source_language}
            Maqsad tili: {target_language}
            Tarjima usuli: {translation_mode}

            Matn:
            {source_text}

            {historical_warning}

            Quyidagi tartibda javob bering:
            1. Tarjima
            2. Muhim so‘zlarning ma'nosi
            3. Zarur bo‘lsa, grammatik izoh
            4. Noaniq yoki bahsli qismlar
            5. Ishonchlilik haqida qisqa eslatma

            Agar foydalanuvchi so‘zma-so‘z tarjimani tanlasa,
            har bir so‘zni alohida ko‘rsat va imkon qadar asl tartibini saqla.
            """

            with st.spinner("JON AI tarjima qilmoqda..."):
                translation = gemini_text(prompt)

            st.session_state["translation_result"] = translation

    translation_result = st.session_state.get("translation_result", "")

    if translation_result:
        st.markdown("---")
        st.subheader("📖 Tarjima natijasi")
        st.markdown(translation_result)

        st.download_button(
            "⬇️ Tarjimani yuklab olish",
            data=translation_result,
            file_name="jon_ai_tarjima.txt",
            mime="text/plain",
            use_container_width=True
        )

# ---------------------- IMAGE PROMPT STUDIO ----------------------

def render_image_prompt():
    st.title("🎨 JON AI — Image Prompt Studio")
    st.caption("Professional rasm yaratish uchun AI prompt generatori")

    subject = st.text_input(
        "Rasm mavzusi",
        placeholder="Masalan: Kelajakdagi aqlli shahar"
    )

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Photorealistic",
            "Cinematic",
            "3D Render",
            "Anime",
            "Digital Art",
            "Fantasy",
            "Minimalist",
            "Cyberpunk",
            "Watercolor",
            "Professional Logo"
        ]
    )

    lighting = st.selectbox(
        "Yoritish",
        [
            "Cinematic lighting",
            "Neon lighting",
            "Soft natural light",
            "Golden hour",
            "Studio lighting",
            "Dramatic lighting"
        ]
    )

    aspect = st.selectbox(
        "Rasm o‘lchami",
        ["1:1", "16:9", "9:16", "4:3", "3:2"]
    )

    details = st.text_area(
        "Qo‘shimcha talablar",
        placeholder="Ranglar, fon, obyektlar va boshqa tafsilotlar..."
    )

    if st.button("✨ Prompt yaratish", use_container_width=True):
        if not subject.strip():
            st.warning("Rasm mavzusini kiriting.")
        else:
            prompt = f"""
            Create a highly detailed professional image-generation prompt.

            Subject: {subject}
            Style: {style}
            Lighting: {lighting}
            Aspect ratio: {aspect}
            Additional requirements: {details}

            Include:
            - Main subject and composition
            - Environment and background
            - Camera angle and perspective
            - Lighting and color palette
            - Materials and textures
            - Visual quality and fine details

            Return:
            1. English image prompt
            2. Uzbek explanation
            3. Negative prompt

            Do not claim that an image has been generated.
            """

            with st.spinner("Professional prompt tayyorlanmoqda..."):
                result = gemini_text(prompt)

            st.session_state["image_prompt_result"] = result

    result = st.session_state.get("image_prompt_result", "")

    if result:
        st.markdown("---")
        st.subheader("🖼️ Tayyor prompt")
        st.markdown(result)

        st.download_button(
            "⬇️ Promptni yuklab olish",
            data=result,
            file_name="jon_ai_image_prompt.txt",
            mime="text/plain",
            use_container_width=True
        )


# ---------------------- EXCEL ANALYSIS ----------------------

def render_excel():
    st.title("📈 JON AI — Excel Intelligence")
    st.caption("Excel va CSV fayllarni tahlil qilish")

    if pd is None:
        st.error("Pandas kutubxonasi o‘rnatilmagan.")
        return

    uploaded_file = st.file_uploader(
        "Excel yoki CSV faylni yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("Fayl muvaffaqiyatli o‘qildi.")

            col1, col2, col3 = st.columns(3)
            col1.metric("Qatorlar", len(df))
            col2.metric("Ustunlar", len(df.columns))
            col3.metric("Bo‘sh qiymatlar", int(df.isna().sum().sum()))

            st.subheader("📋 Jadval ko‘rinishi")
            st.dataframe(df.head(100), use_container_width=True)

            st.subheader("📊 Statistik ma'lumot")
            st.dataframe(
                df.describe(include="all").transpose(),
                use_container_width=True
            )

            st.subheader("🤖 AI tahlili")

            question = st.text_area(
                "Jadval bo‘yicha savolingiz",
                placeholder="Masalan: Savdo natijalaridagi asosiy tendensiyalarni aniqlang."
            )

            if st.button("🔍 Jadvalni AI bilan tahlil qilish", use_container_width=True):
                # Juda katta fayllarni AI ga yuborishda hajmni cheklash
                sample = df.head(30).to_string(index=False)
                columns = ", ".join(str(c) for c in df.columns)

                prompt = f"""
                Quyidagi jadvalni tahlil qil.

                Ustunlar: {columns}

                Jadvalning dastlabki 30 qatori:
                {sample}

                Foydalanuvchi savoli:
                {question or "Umumiy tahlil qil."}

                Quyidagilarni ko‘rib chiq:
                1. Asosiy tendensiyalar
                2. Muhim raqamlar
                3. Ehtimoliy muammolar
                4. Amaliy tavsiyalar

                Faqat jadvalda mavjud ma'lumotlarga asoslan.
                Butun faylni ko‘rmaganingni hisobga ol va xulosalarni
                faqat yuborilgan namunaga tegishli deb belgilagin.
                """

                with st.spinner("Ma'lumotlar tahlil qilinmoqda..."):
                    result = gemini_text(prompt)

                st.session_state["excel_result"] = result

            if st.session_state.get("excel_result"):
                st.markdown("### 📑 Tahlil natijasi")
                st.markdown(st.session_state["excel_result"])

            csv_data = df.to_csv(index=False).encode("utf-8-sig")

            st.download_button(
                "⬇️ CSV formatida yuklab olish",
                data=csv_data,
                file_name="jon_ai_export.csv",
                mime="text/csv",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Faylni o‘qishda xatolik: {e}")


# ---------------------- TASK MANAGER ----------------------

def render_tasks():
    st.title("📝 JON AI — Vazifalar boshqaruvi")
    st.caption("Maqsadlaringiz va kundalik ishlaringizni boshqaring")

    with st.form("add_task_form", clear_on_submit=True):
        task_title = st.text_input(
            "Yangi vazifa",
            placeholder="Masalan: JON AI loyihasini rivojlantirish"
        )

        priority = st.selectbox(
            "Muhimlik darajasi",
            ["Oddiy", "Muhim", "Juda muhim"]
        )

        deadline = st.date_input(
            "Muddat",
            value=date.today()
        )

        submitted = st.form_submit_button(
            "➕ Vazifa qo‘shish",
            use_container_width=True
        )

        if submitted:
            if task_title.strip():
                st.session_state["tasks"].append({
                    "title": task_title.strip(),
                    "priority": priority,
                    "deadline": str(deadline),
                    "done": False
                })
                st.success("Vazifa qo‘shildi.")
            else:
                st.warning("Vazifa nomini kiriting.")

    tasks = st.session_state["tasks"]

    st.markdown("---")
    st.subheader("📌 Vazifalar ro‘yxati")

    if not tasks:
        st.info("Hozircha vazifalar yo‘q. Birinchi vazifangizni qo‘shing.")
        return

    completed = sum(1 for task in tasks if task["done"])
    total = len(tasks)

    st.progress(completed / total)
    st.caption(f"Bajarilgan: {completed} / {total}")

    for index, task in enumerate(tasks):
        col1, col2, col3 = st.columns([5, 1, 1])

        with col1:
            st.markdown(f"**{task['title']}**")
            st.caption(
                f"{task['priority']} · Muddat: {task['deadline']}"
            )

        with col2:
            checked = st.checkbox(
                "Bajarildi",
                value=task["done"],
                key=f"task_done_{index}"
            )
            st.session_state["tasks"][index]["done"] = checked

        with col3:
            if st.button("🗑️", key=f"task_delete_{index}"):
                st.session_state["tasks"].pop(index)
                st.rerun()

    if st.button("🧹 Bajarilgan vazifalarni tozalash"):
        st.session_state["tasks"] = [
            task for task in st.session_state["tasks"]
            if not task["done"]
        ]
        st.rerun()
        
# ---------------------- CAMERA ----------------------

def render_camera():
    st.title("📷 JON AI — Live Camera")
    st.caption("Brauzer kamerasi orqali jonli tasvir")

    st.warning(
        "Kamerani yoqish uchun brauzer ruxsat so‘raydi. "
        "Ruxsat berish yoki bermaslik sizning ixtiyoringizda."
    )

    components.html(
        """
        <div style="
            background:#08152b;
            border:1px solid #36bdf2;
            border-radius:18px;
            padding:15px;
            color:white;
            text-align:center;
            font-family:Arial;
        ">
            <video id="camera" autoplay playsinline
                style="width:100%;max-width:600px;border-radius:14px;
                background:#000;min-height:200px">
            </video>

            <p id="camera-status" style="color:#82eaff">
                Kamera hali yoqilmagan
            </p>

            <button id="start-camera" style="
                padding:12px 22px;border:0;border-radius:10px;
                background:#087eaf;color:white;font-weight:bold;
                margin:8px">
                Kamerani yoqish
            </button>

            <button id="stop-camera" style="
                padding:12px 22px;border:0;border-radius:10px;
                background:#63335f;color:white;font-weight:bold;
                margin:8px">
                Kamerani o‘chirish
            </button>
        </div>

        <script>
        const video = document.getElementById("camera");
        const status = document.getElementById("camera-status");
        const start = document.getElementById("start-camera");
        const stop = document.getElementById("stop-camera");

        let stream = null;

        start.addEventListener("click", async () => {
            try {
                stream = await navigator.mediaDevices.getUserMedia({
                    video: true,
                    audio: false
                });

                video.srcObject = stream;
                status.textContent = "🟢 Kamera ishlayapti";
            } catch (error) {
                status.textContent =
                    "Kameraga ruxsat berilmadi yoki qurilma mavjud emas.";
            }
        });

        stop.addEventListener("click", () => {
            if (stream) {
                stream.getTracks().forEach(track => track.stop());
                stream = null;
            }

            video.srcObject = null;
            status.textContent = "Kamera o‘chirildi";
        });
        </script>
        """,
        height=430,
        scrolling=False
    )

    st.info(
        "Hozirgi versiyada kamera faqat brauzer ichida tasvirni ko‘rsatadi. "
        "JON AI tasvirni avtomatik tahlil qilmaydi."
    )


# ---------------------- CHARACTER SETTINGS ----------------------

def render_character_page():
    st.title("👤 JON AI — Personaj studiyasi")
    st.caption("JON personajining ko‘rinishi va harakatlarini sozlash")

    col1, col2 = st.columns(2)

    with col1:
        st.text_input(
            "Personaj nomi",
            key="character_name"
        )

        st.selectbox(
            "Personaj uslubi",
            [
                "Hologramma",
                "Realistik robot",
                "Multfilm uslubi",
                "Fantastik"
            ],
            key="character_style"
        )

    with col2:
        st.selectbox(
            "Harakat uslubi",
            [
                "Tabiiy harakat",
                "Raqs",
                "Sokin",
                "Energetik"
            ],
            key="character_animation"
        )

        st.selectbox(
            "Kayfiyat",
            [
                "neutral",
                "happy",
                "thinking",
                "serious",
                "greeting"
            ],
            key="jon_mood"
        )

    st.markdown("---")
    st.subheader("✨ Jonli ko‘rinish")
    show_character(height=500)

    st.success(
        "Sozlamalar joriy sessiyada qo‘llanadi. "
        "Doimiy saqlash funksiyasi keyingi versiyada qo‘shilishi mumkin."
    )


# ---------------------- PREMIUM ----------------------

def render_premium():
    st.title("💎 JON AI PREMIUM")
    st.caption("Kengaytirilgan ijodiy va professional imkoniyatlar")

    st.markdown("""
    <div class="motivation-frame">
        <div class="motivation-label">JON AI PREMIUM</div>
        <h1 style="color:#8ceaff">$5 / oy</h1>
        <p style="color:#d2e5ff">
            Rejalashtirilgan obuna narxi
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("🚀 Rejalashtirilgan Premium imkoniyatlar")

    features = [
        "🎨 Yuqori sifatli AI rasm yaratish",
        "🎬 AI video yaratish",
        "✂️ Video montaj va avtomatik montaj",
        "🏢 Professional kompaniya logotiplari",
        "🎵 AI Music Studio",
        "🖼️ Photo to Movie",
        "🎭 Photo to Cartoon",
        "🤖 AI Character Animation",
        "🎞️ Movie Maker",
        "📊 Kengaytirilgan taqdimot vositalari"
    ]

    for feature in features:
        st.markdown(f"- {feature}")

    st.info(
        "Diqqat: $5 narx hozircha rejalashtirilgan. "
        "Haqiqiy obuna, to‘lov tizimi va Premium xizmatlar "
        "hali ulanmagan."
    )

    st.markdown("---")
    st.subheader("📢 Reklama siyosati")

    st.write(
        "JON AI foydalanuvchilarni ortiqcha reklama bilan bezovta qilmasligi "
        "kerak. Reklama soni cheklangan va foydalanuvchi tajribasiga "
        "xalaqit bermaydigan tarzda rejalashtiriladi."
    )

    st.caption(
        "Reklama tizimi va obuna boshqaruvi hali ishlab chiqilmagan."
    )
    
# ---------------------- MAIN APP ROUTER ----------------------

def render_app():
    render_navigation()

    selected_page = st.session_state.get(
        "selected_page",
        PAGES[0]
    )

    if selected_page == "🏠 Bosh sahifa":
        render_home()

    elif selected_page == "🤖 AI Chat":
        render_chat()

    elif selected_page == "🎓 Ta'lim markazi":
        render_education()

    elif selected_page == "📊 PowerPoint":
        render_powerpoint()

    elif selected_page == "🌐 Tarjimon":
        render_translator()

    elif selected_page == "🎨 Rasm prompti":
        render_image_prompt()

    elif selected_page == "📈 Excel tahlil":
        render_excel()

    elif selected_page == "📝 Vazifalar":
        render_tasks()

    elif selected_page == "📷 Kamera":
        render_camera()

    elif selected_page == "👤 Personaj":
        render_character_page()

    elif selected_page == "⚙️ Sozlamalar":
        render_settings()

    elif selected_page == "💎 Premium":
        render_premium()

    else:
        st.warning("Bu sahifa topilmadi.")
        st.session_state["selected_page"] = PAGES[0]


# ---------------------- SETTINGS ----------------------

def render_settings():
    st.title("⚙️ JON AI — System Settings")
    st.caption("API, model va tizim sozlamalari")

    st.subheader("🔑 Gemini API")

    current_key = st.session_state.get("api_key", "")

    api_key = st.text_input(
        "Gemini API kaliti",
        value=current_key,
        type="password",
        placeholder="Gemini API kalitingizni kiriting"
    )

    if st.button("💾 API kalitni saqlash", use_container_width=True):
        st.session_state["api_key"] = api_key.strip()
        st.success("API kalit joriy sessiya uchun saqlandi.")

    st.caption(
        "Xavfsizlik: API kalitni hech kimga yubormang. "
        "Bu oddiy sessiya saqlashi bo‘lib, doimiy xavfsiz saqlash emas."
    )

    st.markdown("---")
    st.subheader("🧠 Gemini modeli")

    model_options = [
        "gemini-3.5-flash-lite",
    ]

    selected_model = st.selectbox(
        "Gemini modeli",
        model_options,
        index=0
    )

    st.session_state["model"] = selected_model
    st.success(f"Tanlangan model: {selected_model}")

    st.markdown("---")
    st.subheader("🔌 API ulanishini tekshirish")

    if st.button("Ulanishni tekshirish", use_container_width=True):
        if not get_api_key():
            st.error("Avval Gemini API kalitini kiriting.")
        else:
            with st.spinner("Gemini bilan bog‘lanilmoqda..."):
                test_result = gemini_text(
                    "Faqat JON AI tizimi muvaffaqiyatli ulangan deb javob ber."
                )

            if "xatolik" in test_result.lower() or "kalit" in test_result.lower():
                st.error(test_result)
            else:
                st.success("Gemini javob qaytardi.")
                st.write(test_result)

    st.markdown("---")
    st.subheader("🧹 Tizim boshqaruvi")

    if st.button("Suhbat tarixini tozalash"):
        st.session_state["messages"] = []
        st.success("Suhbat tarixi tozalandi.")

    st.caption(f"JON AI versiyasi: {APP_VERSION}")                    
# ---------------------- APPLICATION START ----------------------

if __name__ == "__main__":
    render_app()
    
        
