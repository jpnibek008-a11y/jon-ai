
# ============================================================
# JON AI LIVE 9.0 — NEXT GENERATION VIRTUAL INTELLIGENCE
# Streamlit + Gemini AI + Three.js Dual 3D Characters
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

APP_VERSION = "9.0"
DEFAULT_MODEL = "gemini-2.5-flash"

st.set_page_config(
    page_title="JON AI LIVE 9.0",
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
Siz JON AI nomli professional sun'iy intellekt yordamchisiz.
Foydalanuvchiga o'zbek tilida aniq, foydali va tushunarli javob bering.
Kerak bo'lsa ingliz, rus va boshqa tillarda ham javob bera olasiz.
Bilmagan ma'lumotingizni to'qib chiqarmang.
Dasturlash, ta'lim, tarjima, biznes, ijod va tahlilda yordam bering.
Javoblarni tartibli va amaliy shaklda taqdim eting.
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
    "model": DEFAULT_MODEL,
    "intro_enabled": True,
    "page": PAGES[0],
    "copied_prompt": "",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# 3. DESIGN SYSTEM
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;500;600;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

:root {
    --bg: #050817;
    --panel: #0c1430;
    --blue: #26baff;
    --purple: #9b5cff;
    --text: #edf4ff;
}

.stApp {
    background:
        radial-gradient(ellipse at 15% 10%, #17265a 0%, transparent 35%),
        radial-gradient(ellipse at 85% 20%, #25124d 0%, transparent 35%),
        linear-gradient(145deg, #050817, #080d20 60%, #09051c);
    color: var(--text);
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
    font-size: clamp(30px, 5vw, 64px);
    background: linear-gradient(90deg, #ffffff, #43caff, #a879ff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    text-shadow: 0 0 30px #35aaff44;
}

.hero-subtitle {
    color: #9bb8df;
    text-align: center;
    letter-spacing: 4px;
    font-size: 12px;
}

.panel {
    background: linear-gradient(145deg, #111c3bdb, #10102bdd);
    border: 1px solid #4e8cff44;
    border-radius: 18px;
    padding: 20px;
    box-shadow: 0 0 25px #176dff12;
    margin-bottom: 15px;
}

.feature-card {
    background: linear-gradient(145deg, #111c38, #15102e);
    border: 1px solid #438dff44;
    border-radius: 15px;
    padding: 18px;
    min-height: 125px;
    transition: 0.3s;
}

.feature-card:hover {
    border-color: #43caff;
    box-shadow: 0 0 20px #43caff33;
    transform: translateY(-3px);
}

div.stButton > button {
    border-radius: 12px;
    border: 1px solid #438dff77;
    background: linear-gradient(100deg, #123c75, #39206b);
    color: white;
    font-weight: 600;
    min-height: 42px;
}

div.stButton > button:hover {
    border-color: #54d5ff;
    color: white;
    box-shadow: 0 0 15px #36baff55;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #080e23, #100d29);
    border-right: 1px solid #3475b944;
}

[data-testid="stChatMessage"] {
    background: #101a35;
    border: 1px solid #3d6eb533;
    border-radius: 14px;
}

.stTextInput input, .stTextArea textarea {
    background: #0d1730 !important;
    color: white !important;
    border-color: #3475b966 !important;
}

.stSelectbox div[data-baseweb="select"] {
    background: #101a35;
}

hr {
    border-color: #3c69a633;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 4. GEMINI API
# ============================================================

def get_api_key():
    key = str(st.session_state.get("api_key", "")).strip()

    if key:
        return key

    try:
        return str(st.secrets.get("GEMINI_API_KEY", "")).strip()
    except Exception:
        return ""


def get_client():
    key = get_api_key()

    if not key:
        raise ValueError(
            "Gemini API kalitini Sozlamalar bo'limiga kiriting."
        )

    return genai.Client(api_key=key)


def explain_api_error(error):
    message = str(error)
    lower = message.lower()

    if "429" in message or "resource_exhausted" in lower:
        return (
            "Gemini API kvotasi yoki so'rov limiti tugagan. "
            "Biroz kuting yoki Google AI Studio'dagi limitlarni tekshiring."
        )

    if "503" in message or "unavailable" in lower:
        return (
            "Gemini serveri hozir band yoki vaqtincha ishlamayapti. "
            "Birozdan keyin qayta urinib ko'ring."
        )

    if "401" in message or "403" in message or "api key" in lower:
        return "API kalit noto'g'ri yoki uning ruxsatlari yetarli emas."

    if "404" in message or "not found" in lower:
        return "Tanlangan Gemini modeli mavjud emas. Sozlamalardan modelni almashtiring."

    if "timeout" in lower:
        return "Ulanish vaqti tugadi. Internetni tekshirib, qayta urinib ko'ring."

    return f"Gemini bilan ishlashda xatolik: {message[:500]}"


def ask_ai(prompt, extra_context="", history=None):
    """Gemini AI bilan suhbat."""

    try:
        client = get_client()

        model = st.session_state.get("model", DEFAULT_MODEL)
        history = history or []

        conversation = [JON_CORE]

        for item in history[-12:]:
            role = item.get("role", "user")
            content = item.get("content", "")

            if not content:
                continue

            role_name = "Foydalanuvchi" if role == "user" else "JON AI"
            conversation.append(f"{role_name}: {content}")

        if extra_context:
            conversation.append(f"Qo'shimcha ma'lumot: {extra_context}")

        conversation.append(f"Foydalanuvchi: {prompt}")
        conversation.append("JON AI:")

        full_prompt = "\n\n".join(conversation)

        response = client.models.generate_content(
            model=model,
            contents=full_prompt
        )

        answer = getattr(response, "text", None)

        if not answer:
            return "AI javob qaytarmadi. Savolni qayta yuboring."

        return answer.strip()

    except Exception as error:
        return explain_api_error(error)


# ============================================================
# 5. THREE.JS DUAL CHARACTER EXPERIENCE
# ============================================================

def show_character(height=610, cinematic=True):
    name = json.dumps(st.session_state.get("character_name", "JON"))
    style = json.dumps(st.session_state.get("character_style", "Hologramma"))
    intro = "true" if cinematic and st.session_state.get("intro_enabled", True) else "false"

    html = r"""
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* { box-sizing: border-box; }
html, body {
    margin: 0;
    width: 100%;
    height: 100%;
    overflow: hidden;
    background: transparent;
    font-family: Arial, sans-serif;
}
#scene {
    width: 100%;
    height: 100%;
    position: relative;
    overflow: hidden;
    border-radius: 18px;
    background:
      radial-gradient(ellipse at 50% 60%, #142b58aa, transparent 55%),
      radial-gradient(ellipse at 50% 20%, #21134b99, transparent 65%),
      #050817;
}
canvas { display: block; width: 100%; height: 100%; }

.brand {
    position: absolute;
    z-index: 4;
    top: 17px;
    left: 0;
    width: 100%;
    text-align: center;
    pointer-events: none;
}
.brand h1 {
    margin: 0;
    color: white;
    font-size: clamp(22px, 4vw, 37px);
    letter-spacing: 7px;
    text-shadow: 0 0 20px #36caff;
}
.brand span {
    color: #a8c8ef;
    font-size: 9px;
    letter-spacing: 5px;
}
.labels {
    position: absolute;
    bottom: 17px;
    width: 100%;
    display: flex;
    justify-content: space-around;
    z-index: 4;
    pointer-events: none;
}
.label {
    text-align: center;
    color: #c5e8ff;
    font-size: 11px;
    letter-spacing: 3px;
    text-shadow: 0 0 12px #3bbcff;
}
.label small {
    display: block;
    margin-top: 5px;
    color: #7b9dc9;
    font-size: 8px;
    letter-spacing: 2px;
}
.status {
    position: absolute;
    z-index: 4;
    top: 90px;
    left: 50%;
    transform: translateX(-50%);
    color: #5effce;
    font-size: 9px;
    letter-spacing: 3px;
    text-shadow: 0 0 10px #00ffaa;
}
#intro {
    position: absolute;
    z-index: 5;
    inset: 0;
    background: radial-gradient(circle, #173b75 0%, #090d22 55%, #02040c 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-size: 15px;
    letter-spacing: 6px;
    transition: opacity 1.5s ease;
    pointer-events: none;
}
#intro.hide { opacity: 0; }
</style>
</head>
<body>
<div id="scene">
    <div class="brand">
        <h1>JON AI</h1>
        <span>NEXT GENERATION VIRTUAL INTELLIGENCE</span>
    </div>
    <div class="status" id="status">● SYSTEM ONLINE</div>
    <div class="labels">
        <div class="label">HUMAN<small>VIRTUAL INTELLIGENCE</small></div>
        <div class="label">ANDROID<small>ADVANCED TECHNOLOGY</small></div>
    </div>
    <div id="intro">INITIALIZING JON AI</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
(function() {
    const sceneBox = document.getElementById("scene");
    const intro = document.getElementById("intro");
    const status = document.getElementById("status");

    if (!window.THREE) {
        intro.innerText = "3D ENGINE LOAD FAILED";
        return;
    }

    const THREE = window.THREE;
    const scene = new THREE.Scene();
    scene.fog = new THREE.FogExp2(0x050817, 0.035);

    const camera = new THREE.PerspectiveCamera(
        42,
        sceneBox.clientWidth / sceneBox.clientHeight,
        0.1,
        100
    );
    camera.position.set(0, 1.1, 10.5);

    const renderer = new THREE.WebGLRenderer({
        antialias: true,
        alpha: true
    });

    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.5));
    renderer.setSize(sceneBox.clientWidth, sceneBox.clientHeight);
    renderer.outputEncoding = THREE.sRGBEncoding;
    sceneBox.appendChild(renderer.domElement);

    scene.add(new THREE.AmbientLight(0x7899ff, 1.5));

    const keyLight = new THREE.PointLight(0x38bfff, 2.5, 30);
    keyLight.position.set(-4, 5, 5);
    scene.add(keyLight);

    const purpleLight = new THREE.PointLight(0x974bff, 2.2, 25);
    purpleLight.position.set(4, 3, 3);
    scene.add(purpleLight);

    const floorLight = new THREE.PointLight(0x176bff, 1.8, 20);
    floorLight.position.set(0, -3, 4);
    scene.add(floorLight);

    const material = (color, metalness=0.5, roughness=0.3) =>
        new THREE.MeshStandardMaterial({
            color,
            metalness,
            roughness,
            emissive: color,
            emissiveIntensity: 0.16
        });

    const blue = material(0x167bd1, 0.7, 0.25);
    const cyan = material(0x55dfff, 0.6, 0.2);
    const dark = material(0x101d38, 0.8, 0.28);
    const skin = material(0xc58e78, 0.1, 0.55);
    const white = material(0xc9eaff, 0.5, 0.2);
    const purple = material(0x8c48e8, 0.7, 0.2);

    const glowMaterial = new THREE.MeshBasicMaterial({
        color: 0x44d8ff,
        transparent: true,
        opacity: 0.8
    });

    function sphere(parent, x, y, z, sx, sy, sz, mat) {
        const mesh = new THREE.Mesh(
            new THREE.SphereGeometry(1, 24, 18),
            mat
        );
        mesh.position.set(x, y, z);
        mesh.scale.set(sx, sy, sz);
        parent.add(mesh);
        return mesh;
    }

    function cylinder(parent, x, y, z, rt, rb, h, mat) {
        const mesh = new THREE.Mesh(
            new THREE.CylinderGeometry(rt, rb, h, 20),
            mat
        );
        mesh.position.set(x, y, z);
        parent.add(mesh);
        return mesh;
    }

    function ring(parent, radius, tube, color, x, y, z, rotationX) {
        const mesh = new THREE.Mesh(
            new THREE.TorusGeometry(radius, tube, 10, 90),
            new THREE.MeshBasicMaterial({
                color,
                transparent: true,
                opacity: 0.75
            })
        );
        mesh.position.set(x, y, z);
        mesh.rotation.x = rotationX;
        parent.add(mesh);
        return mesh;
    }

    function makeHuman() {
        const root = new THREE.Group();
        root.position.x = -2.05;
        scene.add(root);

        const body = new THREE.Group();
        root.add(body);

        // Legs
        cylinder(body, -0.22, -1.05, 0, 0.13, 0.17, 0.9, dark);
        cylinder(body, 0.22, -1.05, 0, 0.13, 0.17, 0.9, dark);
        sphere(body, -0.22, -1.48, 0.1, 0.19, 0.12, 0.3, blue);
        sphere(body, 0.22, -1.48, 0.1, 0.19, 0.12, 0.3, blue);

        // Torso and shoulders
        sphere(body, 0, -0.05, 0, 0.57, 0.78, 0.32, blue);
        sphere(body, 0, -0.08, 0.29, 0.38, 0.48, 0.1, dark);

        // Neck
        cylinder(body, 0, 0.72, 0, 0.15, 0.17, 0.32, skin);

        // Head
        const head = new THREE.Group();
        head.position.y = 1.35;
        body.add(head);

        sphere(head, 0, 0, 0, 0.43, 0.55, 0.36, skin);
        sphere(head, 0, 0.39, -0.015, 0.39, 0.19, 0.34, dark);

        // Eyes
        const eyeMat = new THREE.MeshBasicMaterial({color: 0x8beaff});
        const eyeL = sphere(head, -0.16, 0.08, 0.32, 0.055, 0.035, 0.025, eyeMat);
        const eyeR = sphere(head, 0.16, 0.08, 0.32, 0.055, 0.035, 0.025, eyeMat);

        // Mouth
        const mouth = sphere(head, 0, -0.22, 0.33, 0.11, 0.018, 0.02, dark);

        // Arms
        const armL = new THREE.Group();
        const armR = new THREE.Group();
        armL.position.set(-0.57, 0.36, 0);
        armR.position.set(0.57, 0.36, 0);
        body.add(armL, armR);

        cylinder(armL, 0, -0.38, 0, 0.13, 0.17, 0.7, skin);
        cylinder(armR, 0, -0.38, 0, 0.13, 0.17, 0.7, skin);
        sphere(armL, 0, -0.78, 0, 0.15, 0.16, 0.15, skin);
        sphere(armR, 0, -0.78, 0, 0.15, 0.16, 0.15, skin);

        ring(root, 0.9, 0.018, 0x45cfff, 0, -1.7, 0, 0);
        ring(root, 0.75, 0.012, 0x8e55ff, 0, -1.7, 0, 0);

        return {root, body, head, mouth, eyeL, eyeR, armL, armR};
    }

    function makeAndroid() {
        const root = new THREE.Group();
        root.position.x = 2.05;
        scene.add(root);

        const body = new THREE.Group();
        root.add(body);

        // Legs
        cylinder(body, -0.22, -1.05, 0, 0.15, 0.18, 0.9, dark);
        cylinder(body, 0.22, -1.05, 0, 0.15, 0.18, 0.9, dark);

        sphere(body, -0.22, -1.48, 0.12, 0.2, 0.12, 0.3, blue);
        sphere(body, 0.22, -1.48, 0.12, 0.2, 0.12, 0.3, blue);

        // Torso
        sphere(body, 0, -0.05, 0, 0.58, 0.77, 0.34, dark);
        sphere(body, 0, -0.1, 0.31, 0.34, 0.42, 0.1, blue);
        sphere(body, 0, -0.1, 0.42, 0.15, 0.16, 0.035, cyan);

        // Neck and head
        cylinder(body, 0, 0.75, 0, 0.17, 0.17, 0.28, blue);

        const head = new THREE.Group();
        head.position.y = 1.36;
        body.add(head);

        sphere(head, 0, 0, 0, 0.45, 0.48, 0.38, dark);
        sphere(head, 0, 0.02, 0.29, 0.36, 0.25, 0.12, blue);

        // Face visor
        sphere(head, 0, 0.02, 0.39, 0.29, 0.13, 0.035, cyan);

        const eyeMat = new THREE.MeshBasicMaterial({
            color: 0xb9f8ff
        });
        const eyeL = sphere(head, -0.14, 0.04, 0.43, 0.06, 0.035, 0.025, eyeMat);
        const eyeR = sphere(head, 0.14, 0.04, 0.43, 0.06, 0.035, 0.025, eyeMat);
        const mouth = sphere(head, 0, -0.17, 0.41, 0.12, 0.018, 0.02, purple);

        // Side ear lights
        sphere(head, -0.43, 0, 0, 0.07, 0.15, 0.13, cyan);
        sphere(head, 0.43, 0, 0, 0.07, 0.15, 0.13, cyan);

        // Arms
        const armL = new THREE.Group();
        const armR = new THREE.Group();
        armL.position.set(-0.58, 0.36, 0);
        armR.position.set(0.58, 0.36, 0);
        body.add(armL, armR);

        cylinder(armL, 0, -0.38, 0, 0.16, 0.19, 0.72, blue);
        cylinder(armR, 0, -0.38, 0, 0.16, 0.19, 0.72, blue);
        sphere(armL, 0, -0.78, 0, 0.17, 0.17, 0.17, dark);
        sphere(armR, 0, -0.78, 0, 0.17, 0.17, 0.17, dark);

        ring(root, 1.0, 0.022, 0x36cfff, 0, -1.7, 0, 0);
        ring(root, 0.77, 0.014, 0xa55cff, 0, -1.7, 0, 0);

        return {root, body, head, mouth, eyeL, eyeR, armL, armR};
    }

    const human = makeHuman();
    const android = makeAndroid();

    // Floating energy rings
    const energyRings = [];
    for (let i = 0; i < 5; i++) {
        const r = ring(
            scene,
            2.3 + i * 0.24,
            0.008,
            i % 2 ? 0x9655ff : 0x36cfff,
            0,
            -0.1,
            -1.2,
            Math.PI / 2
        );
        energyRings.push(r);
    }

    // Portal
    const portal = new THREE.Group();
    portal.position.set(0, 0.1, -1.5);
    scene.add(portal);

    const portalRing = new THREE.Mesh(
        new THREE.TorusGeometry(3.15, 0.035, 12, 120),
        new THREE.MeshBasicMaterial({
            color: 0x4d9fff,
            transparent: true,
            opacity: 0.7
        })
    );
    portal.add(portalRing);

    // Stars and particles
    const particleCount = window.innerWidth < 600 ? 350 : 650;
    const positions = new Float32Array(particleCount * 3);

    for (let i = 0; i < particleCount; i++) {
        positions[i * 3] = (Math.random() - 0.5) * 15;
        positions[i * 3 + 1] = (Math.random() - 0.5) * 9;
        positions[i * 3 + 2] = (Math.random() - 0.5) * 9;
    }

    const particleGeometry = new THREE.BufferGeometry();
    particleGeometry.setAttribute(
        "position",
        new THREE.BufferAttribute(positions, 3)
    );

    const particleMaterial = new THREE.PointsMaterial({
        color: 0x72caff,
        size: 0.035,
        transparent: true,
        opacity: 0.8
    });

    const particles = new THREE.Points(
        particleGeometry,
        particleMaterial
    );
    scene.add(particles);

    // Energy arcs
    const arcs = [];

    for (let i = 0; i < 9; i++) {
        const points = [];
        const start = new THREE.Vector3(
            (Math.random() - 0.5) * 5,
            (Math.random() - 0.5) * 3,
            0.2
        );

        let current = start.clone();

        for (let j = 0; j < 6; j++) {
            current = current.clone().add(
                new THREE.Vector3(
                    (Math.random() - 0.5) * 0.55,
                    (Math.random() - 0.5) * 0.55,
                    0
                )
            );
            points.push(current.clone());
        }

        const geometry = new THREE.BufferGeometry().setFromPoints(points);
        const line = new THREE.Line(
            geometry,
            new THREE.LineBasicMaterial({
                color: i % 2 ? 0x9b5cff : 0x42dfff,
                transparent: true,
                opacity: 0.0
            })
        );

        scene.add(line);
            arcs.push(line);
    }

    // Floor platform
    const platform = new THREE.Mesh(
        new THREE.CylinderGeometry(3.6, 3.6, 0.08, 80),
        new THREE.MeshBasicMaterial({
            color: 0x0c2852,
            transparent: true,
            opacity: 0.42
        })
    );
    platform.position.y = -1.72;
    platform.position.z = -0.3;
    scene.add(platform);

    const clock = new THREE.Clock();
    let frame = 0;
    let isSpeaking = false;

    // Future voice integration
    window.JONCharacter = {
        speak: function(value) {
            isSpeaking = !!value;
            status.innerText = isSpeaking
                ? "● JON AI IS SPEAKING"
                : "● SYSTEM ONLINE";
        }
    };

    function animate() {
        requestAnimationFrame(animate);

        const t = clock.getElapsedTime();
        frame++;

        [human, android].forEach((character, index) => {
            const phase = index * 1.3;

            character.root.position.y =
                Math.sin(t * 1.5 + phase) * 0.08;

            character.root.rotation.y =
                Math.sin(t * 0.45 + phase) * 0.08;

            character.head.rotation.y =
                Math.sin(t * 0.8 + phase) * 0.08;

            character.body.rotation.z =
                Math.sin(t * 1.1 + phase) * 0.025;

            character.armL.rotation.z =
                Math.sin(t * 1.4 + phase) * 0.08;

            character.armR.rotation.z =
                -Math.sin(t * 1.4 + phase) * 0.08;

            const mouthMove = isSpeaking
                ? 0.04 + Math.abs(Math.sin(t * 12)) * 0.08
                : 0.018;

            character.mouth.scale.y = mouthMove;

            const blink = Math.sin(t * 1.1 + phase * 4) > 0.995
                ? 0.15 : 1;

            character.eyeL.scale.y = 0.035 * blink;
            character.eyeR.scale.y = 0.035 * blink;
        });

        energyRings.forEach((r, i) => {
            r.rotation.z = t * (0.12 + i * 0.035);
            r.material.opacity =
                0.35 + Math.sin(t * 2 + i) * 0.2;
        });

        portalRing.rotation.z = t * 0.08;
        portal.scale.setScalar(
            1 + Math.sin(t * 0.7) * 0.025
        );

        particles.rotation.y = t * 0.012;
        particles.rotation.z = Math.sin(t * 0.1) * 0.02;

        arcs.forEach((arc, i) => {
            arc.material.opacity =
                Math.max(0, Math.sin(t * 3 + i * 2) * 0.45);
        });

        keyLight.intensity = 2.2 + Math.sin(t * 2) * 0.4;
        purpleLight.intensity = 2 + Math.cos(t * 1.6) * 0.35;

        renderer.render(scene, camera);
    }

    function resize() {
        const w = sceneBox.clientWidth;
        const h = sceneBox.clientHeight;

        if (!w || !h) return;

        camera.aspect = w / h;
        camera.updateProjectionMatrix();
        renderer.setSize(w, h);
    }

    window.addEventListener("resize", resize);

    if (""" + intro + r""") {
        setTimeout(() => intro.classList.add("hide"), 800);
        setTimeout(() => intro.remove(), 2400);
    } else {
        intro.remove();
    }

    animate();
})();
</script>
</body>
</html>
"""

    components.html(html, height=height, scrolling=False)


# ============================================================
# 6. HOME PAGE
# ============================================================

def render_home():
    st.markdown(
        '<div class="hero-title">JON AI</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="hero-subtitle">NEXT GENERATION VIRTUAL INTELLIGENCE</div>',
        unsafe_allow_html=True
    )

    show_character(height=620, cinematic=True)

    st.markdown("### ✨ JON AI imkoniyatlari")

    columns = st.columns(3)

    features = [
        ("🤖", "AI Chat", "Sun'iy intellekt bilan muloqot"),
        ("🎓", "Ta'lim markazi", "IELTS, Multilevel va grammatika"),
        ("📊", "PowerPoint", "Taqdimotlar tayyorlash"),
        ("🌐", "Tarjimon", "Ko'p tilli tarjima"),
        ("📈", "Excel tahlil", "Jadval va ma'lumotlar tahlili"),
        ("🎨", "Creative Studio", "Professional promptlar")
    ]

    for i, feature in enumerate(features):
        with columns[i % 3]:
            st.markdown(
                f"""
                <div class="feature-card">
                    <div style="font-size:28px">{feature[0]}</div>
                    <b>{feature[1]}</b>
                    <div style="color:#9bb8df;font-size:13px;margin-top:8px">
                        {feature[2]}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.write("")

    if st.button(
        "🤖 JON AI bilan suhbatni boshlash",
        use_container_width=True
    ):
        st.session_state.page = "🤖 AI Chat"
        st.rerun()


# ============================================================
# 7. AI CHAT
# ============================================================

def render_chat():
    st.title("🤖 JON AI — AI Chat")
    st.caption("Sun'iy intellekt bilan jonli muloqot")

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    prompt = st.chat_input("JON AI'ga xabar yozing...")

    if prompt:
        history = st.session_state.messages.copy()

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI javob tayyorlamoqda..."):
                answer = ask_ai(prompt, history=history)
                st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })


# ============================================================
# 8. EDUCATION CENTER
# ============================================================

def render_education():
    st.title("🎓 JON AI — Ta'lim markazi")

    topic = st.selectbox(
        "Yo'nalish",
        [
            "Ingliz tili", "IELTS", "Multilevel",
            "Grammatika", "Lug'at", "Speaking",
            "Writing", "Reading", "Listening", "Forex"
        ]
    )

    mode = st.selectbox(
        "O'rganish usuli",
        [
            "Mavzuni tushuntirish",
            "Test tuzish",
            "Mashq bajarish",
            "Javoblarni tekshirish",
            "Lug'at tuzish",
            "Dars rejasi"
        ]
    )

    level = st.selectbox(
        "Bilim darajasi",
        [
            "Beginner", "Elementary", "Pre-Intermediate",
            "Intermediate", "Upper-Intermediate", "Advanced"
        ]
    )

    question = st.text_area("Mavzu yoki savolingiz")

    if st.button("🎓 Darsni boshlash", use_container_width=True):
        if not question.strip():
            st.warning("Avval mavzu yoki savol kiriting.")
        else:
            prompt = f"""
            Yo'nalish: {topic}
            O'rganish usuli: {mode}
            Daraja: {level}
            Foydalanuvchi so'rovi: {question}
            Mavzuni bosqichma-bosqich tushuntiring.
            Misollar va amaliy mashqlar bering.
            """

            with st.spinner("Dars tayyorlanmoqda..."):
                st.markdown(ask_ai(prompt))


# ============================================================
# 9. POWERPOINT
# ============================================================

def render_powerpoint():
    st.title("📊 JON AI — PowerPoint Studio")

    topic = st.text_input("Taqdimot mavzusi")
    slide_count = st.slider("Slaydlar soni", 3, 20, 8)
    language = st.selectbox("Taqdimot tili", ["O'zbek", "Ingliz", "Rus"])

    if st.button("✨ Taqdimot yaratish", use_container_width=True):
        if not topic.strip():
            st.warning("Mavzuni kiriting.")
            return

        prompt = f"""
        {topic} mavzusida {slide_count} slaydli taqdimot tayyorla.
        Til: {language}.
        Har bir slayd:
        SLIDE 1: Sarlavha
        BODY:
        - Birinchi fikr
        - Ikkinchi fikr
        """

        with st.spinner("Slaydlar tayyorlanmoqda..."):
            result = ask_ai(prompt)

        st.markdown(result)

        presentation = Presentation()

        for slide_text in re.split(
            r"(?=SLIDE\s+\d+)",
            result,
            flags=re.IGNORECASE
        ):
            slide_text = slide_text.strip()

            if not slide_text:
                continue

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[1]
            )

            lines = slide_text.splitlines()
            slide.shapes.title.text = lines[0][:150]

            body = slide.placeholders[1].text_frame
            body.clear()

            for line in lines[1:]:
                line = re.sub(
                    r"^(TITLE:|BODY:)",
                    "",
                    line
                ).strip()

                if line:
                    paragraph = body.add_paragraph()
                    paragraph.text = line.lstrip("-• ")

        output = BytesIO()
        presentation.save(output)
        output.seek(0)

        st.download_button(
            "📥 PowerPoint yuklab olish",
            data=output,
            file_name="JON_AI_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )


# ============================================================
# 10. TRANSLATOR
# ============================================================

def render_translator():
    st.title("🌐 JON AI — Tarjimon")

    source = st.selectbox(
        "Qaysi tildan?",
        ["Avtomatik aniqlash", "O'zbek", "Ingliz", "Rus", "Turk", "Arab"]
    )

    target = st.selectbox(
        "Qaysi tilga?",
        ["O'zbek", "Ingliz", "Rus", "Turk", "Arab", "Nemis", "Fransuz"]
    )

    text_input = st.text_area("Tarjima qilinadigan matn", height=180)

    if st.button("🌐 Tarjima qilish"):
        if not text_input.strip():
            st.warning("Matn kiriting.")
        else:
            prompt = f"""
            {source} tilidan {target} tiliga tarjima qiling.
            Matn: {text_input}
            """

            with st.spinner("Tarjima qilinmoqda..."):
                result = ask_ai(prompt)

            st.markdown(result)

            st.download_button(
                "📥 Tarjimani yuklab olish",
                result,
                file_name="jon_ai_translation.txt"
            )


# ============================================================
# 11. IMAGE PROMPT
# ============================================================

def render_image_prompt():
    st.title("🎨 JON AI — Image Prompt Studio")

    subject = st.text_input("Rasm mavzusi")

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Realistic", "Cinematic", "3D Render",
            "Anime", "Cyberpunk", "Fantasy",
            "Minimalist", "Professional Branding"
        ]
    )

    ratio = st.selectbox(
        "Format",
        ["1:1", "16:9", "9:16", "4:5", "3:2"]
    )

    details = st.text_area("Qo'shimcha talablar")

    if st.button("✨ Prompt yaratish"):
        if not subject.strip():
            st.warning("Rasm mavzusini kiriting.")
        else:
            prompt = f"""
            Create a professional image generation prompt.
            Subject: {subject}
            Style: {style}
            Aspect ratio: {ratio}
            Additional requirements: {details}
            Include lighting, composition, camera angle and quality.
            Return the prompt in English.
            """

            with st.spinner("Prompt tayyorlanmoqda..."):
                result = ask_ai(prompt)

            st.code(result)

            st.download_button(
                "📥 Promptni yuklab olish",
                result,
                file_name="jon_ai_image_prompt.txt"
            )


# ============================================================
# 12. EXCEL
# ============================================================

def render_excel():
    st.title("📈 JON AI — Excel Analysis")

    uploaded = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded:
        try:
            if uploaded.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded)
            else:
                df = pd.read_excel(uploaded)

            st.dataframe(df, use_container_width=True)

            c1, c2, c3 = st.columns(3)
            c1.metric("Qatorlar", len(df))
            c2.metric("Ustunlar", len(df.columns))
            c3.metric("Bo'sh kataklar", int(df.isna().sum().sum()))

            st.dataframe(
                df.describe(include="all").fillna(""),
                use_container_width=True
            )

            if st.button("🤖 AI orqali tahlil qilish"):
                prompt = (
                    "Jadvalni tahlil qiling, tendensiyalar va xulosalar bering:\n"
                    + df.head(30).to_csv(index=False)
                )

                with st.spinner("Tahlil qilinmoqda..."):
                    st.markdown(ask_ai(prompt))

            st.download_button(
                "📥 CSV yuklab olish",
                df.to_csv(index=False).encode("utf-8-sig"),
                file_name="jon_ai_analysis.csv"
            )

        except Exception as error:
            st.error(f"Xatolik: {error}")


# ============================================================
# 13. TASKS
# ============================================================

def render_tasks():
    st.title("📝 JON AI — Vazifalar")

    task_text = st.text_input("Yangi vazifa")

    if st.button("➕ Vazifa qo'shish"):
        if task_text.strip():
            st.session_state.tasks.append({
                "text": task_text.strip(),
                "done": False
            })
            st.rerun()

    for index, task in enumerate(st.session_state.tasks):
        col1, col2 = st.columns([5, 1])

        with col1:
            checked = st.checkbox(
                task["text"],
                value=task["done"],
                key=f"task_{index}"
            )
            st.session_state.tasks[index]["done"] = checked

        with col2:
            if st.button("🗑", key=f"delete_{index}"):
                st.session_state.tasks.pop(index)
                st.rerun()

    completed = sum(
        1 for task in st.session_state.tasks if task["done"]
    )

    if st.session_state.tasks:
        st.progress(completed / len(st.session_state.tasks))


# ============================================================
# 14. CAMERA
# ============================================================

def render_camera():
    st.title("📷 JON AI — Kamera")

    picture = st.camera_input("Rasmga olish")

    if picture:
        st.image(picture, use_container_width=True)

        st.download_button(
            "📥 Rasmni yuklab olish",
            picture.getvalue(),
            file_name="jon_ai_camera.jpg",
            mime="image/jpeg"
        )


# ============================================================
# 15. CHARACTER STUDIO
# ============================================================

def render_character():
    st.title("👤 JON AI — Character Studio")

    name = st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    styles = [
        "Hologramma",
        "Realistik robot",
        "Multfilm uslubi",
        "Fantastik",
        "Virtual inson"
    ]

    style = st.selectbox(
        "Personaj uslubi",
        styles,
        index=styles.index(st.session_state.character_style)
        if st.session_state.character_style in styles else 0
    )

    if st.button("💾 Saqlash"):
        st.session_state.character_name = name.strip() or "JON"
        st.session_state.character_style = style
        st.success("Sozlamalar saqlandi.")

    show_character(height=580, cinematic=False)

    st.caption(
        "Haqiqiy yuz mimikasi uchun riglangan GLB/GLTF 3D model "
        "va animatsiya tizimi kerak bo'ladi."
    )


# ============================================================
# 16. SETTINGS
# ============================================================

def render_settings():
    st.title("⚙️ JON AI — Sozlamalar")

    api_key = st.text_input(
        "Gemini API kaliti",
        value=st.session_state.api_key,
        type="password"
    )

    if st.button("🔐 API kalitni saqlash"):
        st.session_state.api_key = api_key.strip()
        st.success("API kalit joriy sessiyada saqlandi.")

    st.caption(
        "Doimiy foydalanish uchun Streamlit Secrets'dan foydalaning."
    )

    models = [
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash"
    ]

    current = st.session_state.get("model", DEFAULT_MODEL)

    st.session_state.model = st.selectbox(
        "Gemini modeli",
        models,
        index=models.index(current) if current in models else 0
    )

    st.session_state.intro_enabled = st.toggle(
        "Kirish animatsiyasi",
        value=st.session_state.intro_enabled
    )

    st.write("JON AI versiyasi:", APP_VERSION)


# ============================================================
# 17. PREMIUM
# ============================================================

def render_premium():
    st.title("💎 JON AI — Premium Studio")

    features = [
        ("🎬", "AI Video Studio", "AI video yaratish va montaj"),
        ("🖼️", "AI Image Studio", "Professional tasvirlar"),
        ("🎵", "AI Music Studio", "Original musiqa"),
        ("👤", "AI Character Animation", "3D personaj animatsiyasi"),
        ("🎥", "Movie Maker", "Ssenariy asosida video"),
        ("🏢", "Branding Studio", "Logo va brending"),
        ("🎞️", "Photo to Movie", "Rasmlardan video"),
        ("🎨", "Photo to Cartoon", "Suratlarni multfilmga aylantirish"),
        ("🎙️", "Voice Assistant", "Ovozli yordamchi")
    ]

    for icon, title, description in features:
        st.markdown(
            f"""
            <div class="feature-card">
                <span style="font-size:25px">{icon}</span>
                <b style="margin-left:10px">{title}</b>
                <p style="color:#9bb8df">{description}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.info("Premium funksiyalar bosqichma-bosqich integratsiya qilinadi.")


# ============================================================
# 18. SIDEBAR
# ============================================================

def render_sidebar():
    with st.sidebar:
        st.markdown(
            """
            <div style="text-align:center;padding:15px 0">
                <div style="font-size:38px">💎</div>
                <div style="font-family:Orbitron;font-size:21px;color:#64caff">
                    JON AI
                </div>
                <div style="font-size:10px;letter-spacing:3px;color:#8ca9d1">
                    LIVE INTELLIGENCE
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("---")

        selected = st.radio(
            "Navigatsiya",
            PAGES,
            index=PAGES.index(st.session_state.page)
            if st.session_state.page in PAGES else 0
        )

        st.session_state.page = selected

        st.markdown("---")
        st.caption(f"JON AI LIVE {APP_VERSION}")


# ============================================================
# 19. MAIN
# ============================================================

def main():
    render_sidebar()

    page = st.session_state.page

    routes = {
        "🏠 Bosh sahifa": render_home,
        "🤖 AI Chat": render_chat,
        "🎓 Ta'lim markazi": render_education,
        "📊 PowerPoint": render_powerpoint,
        "🌐 Tarjimon": render_translator,
        "🎨 Rasm prompti": render_image_prompt,
            

    elif page == "📈 Excel tahlil":
        render_excel()

    elif page == "📝 Vazifalar":
        render_tasks()

    elif page == "📷 Kamera":
        render_camera()

    elif page == "👤 Personaj":
        render_character()

    elif page == "⚙️ Sozlamalar":
        render_settings()

    elif page == "💎 Premium":
        render_premium()
    
# ============================================================
# 20. ILOVANI ISHGA TUSHIRISH
# ============================================================

if __name__ == "__main__":
    main()
