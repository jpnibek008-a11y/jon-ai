
import json
import re
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from pptx import Presentation

# ============================================================
# JON AI LIVE 5.1 — BIR FAYLLI UMUMIY VERSIYA
# ============================================================

st.set_page_config(
    page_title="JON AI LIVE",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "5.1"
MODEL = "gemini-2.5-flash"

JON_CORE = """
Sen JON AI nomli aqlli virtual yordamchisan.
Foydalanuvchiga o'zbek tilida ravon, aniq va hurmat bilan javob ber.
Ta'lim, tarjima, dasturlash, tahlil va ijodiy vazifalarda yordam ber.
Agar ishonching komil bo'lmasa, buni ochiq ayt. Ma'lumotni to'qib chiqarma.
"""

# -------------------- SESSION STATE --------------------

DEFAULTS = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "api_key": "",
    "copied_prompt": "",
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# -------------------- AI ENGINE --------------------

def get_api_key():
    key = st.session_state.get("api_key", "").strip()

    if key:
        return key

    try:
        return st.secrets.get("GEMINI_API_KEY", "").strip()
    except Exception:
        return ""


def ask_ai(prompt, extra_context="", history=None):
    api_key = get_api_key()

    if not api_key:
        return (
            "Gemini API kaliti topilmadi. "
            "Sozlamalar bo'limidan API kalitini kiriting."
        )

    try:
        client = genai.Client(api_key=api_key)

        history_lines = []

        for item in (history or [])[-12:]:
            role = (
                "Foydalanuvchi"
                if item.get("role") == "user"
                else "JON AI"
            )

            history_lines.append(
                f"{role}: {item.get('content', '')}"
            )

        contents = (
            JON_CORE
            + "\n\n"
            + (
                "\n".join(history_lines) + "\n\n"
                if history_lines else ""
            )
            + str(extra_context)
            + "\n\nFoydalanuvchi: "
            + str(prompt)
        )

        response = client.models.generate_content(
            model=MODEL,
            contents=contents,
        )

        return response.text or "AI javob qaytarmadi."

    except Exception as exc:
        return f"Gemini bilan bog'lanishda xatolik: {exc}"


# -------------------- THREE.JS 3D PERSONAJ --------------------

def show_character(speaking=False, height=520):

    name = str(st.session_state.character_name)
    style = st.session_state.character_style

    style_map = {
        "Hologramma": "hologram",
        "Robot": "robot",
        "Virtual inson": "human",
        "Multfilm uslubi": "cartoon",
        "Fantastik": "fantasy",
    }

    visual_style = style_map.get(style, "hologram")
    safe_name = json.dumps(name)

    html = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<style>
html, body {
    margin: 0;
    width: 100%;
    height: 100%;
    overflow: hidden;
    background: radial-gradient(
        ellipse at 50% 35%,
        #142b52 0%,
        #080d20 58%,
        #030611 100%
    );
    font-family: Arial, sans-serif;
}

#stage {
    position: relative;
    width: 100%;
    height: 100%;
    min-height: 420px;
}

#scene {
    width: 100%;
    height: 100%;
    display: block;
}

#hud {
    position: absolute;
    top: 15px;
    left: 18px;
    right: 18px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #8ff7ff;
    font-size: 12px;
    letter-spacing: 2px;
    pointer-events: none;
    text-shadow: 0 0 10px #00d9ff;
}

#name {
    position: absolute;
    bottom: 42px;
    left: 0;
    right: 0;
    text-align: center;
    color: #d9fbff;
    font-size: 25px;
    font-weight: bold;
    letter-spacing: 8px;
    text-shadow: 0 0 18px #00cfff;
}

#status {
    position: absolute;
    bottom: 15px;
    left: 0;
    right: 0;
    text-align: center;
    color: #75e8ff;
    font-size: 11px;
    letter-spacing: 3px;
}

#fallback {
    display: none;
    position: absolute;
    inset: 0;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    color: #b8f8ff;
    text-align: center;
    padding: 20px;
}
</style>
</head>

<body>

<div id="stage">

    <div id="hud">
        <span>JON AI LIVE</span>
        <span>● ONLINE</span>
    </div>

    <canvas id="scene"></canvas>

    <div id="name"></div>

    <div id="status">
        VIRTUAL INTELLIGENCE
    </div>

    <div id="fallback">
        <div style="font-size:90px">💎</div>
        <h2>JON AI</h2>
        <p>3D sahna yuklanmadi. Internet aloqasini tekshiring.</p>
    </div>

</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>

<script>
(() => {

    const stage = document.getElementById('stage');
    const canvas = document.getElementById('scene');
    const nameNode = document.getElementById('name');
    const statusNode = document.getElementById('status');
    const fallback = document.getElementById('fallback');

    const characterName = __NAME__;
    const visualStyle = "__STYLE__";
    const isSpeaking = __SPEAKING__;

    nameNode.textContent = characterName;

    if (!window.THREE) {
        fallback.style.display = 'flex';
        canvas.style.display = 'none';
        return;
    }

    const THREE = window.THREE;

    const scene = new THREE.Scene();

    scene.fog = new THREE.Fog(
        0x071126,
        8,
        22
    );

    const camera = new THREE.PerspectiveCamera(
        38,
        stage.clientWidth / stage.clientHeight,
        0.1,
        100
    );

    camera.position.set(0, 0.15, 8.2);

    const renderer = new THREE.WebGLRenderer({
        canvas: canvas,
        antialias: true,
        alpha: true
    });

    renderer.setPixelRatio(
        Math.min(window.devicePixelRatio || 1, 2)
    );

    renderer.setSize(
        stage.clientWidth,
        stage.clientHeight
    );

    renderer.outputEncoding = THREE.sRGBEncoding;

    const colors = {
        hologram: 0x31eaff,
        robot: 0xa8c9df,
        human: 0xffc49a,
        cartoon: 0xffb7e8,
        fantasy: 0xb99aff
    };

    const accent = colors[visualStyle] || colors.hologram;

    scene.add(
        new THREE.AmbientLight(0x9bcfff, 1.2)
    );

    const keyLight = new THREE.PointLight(
        accent,
        2.6,
        18
    );

    keyLight.position.set(3, 3, 5);
    scene.add(keyLight);

    const fillLight = new THREE.PointLight(
        0x287bff,
        1.7,
        16
    );

    fillLight.position.set(-3, -1, 3);
    scene.add(fillLight);

    const character = new THREE.Group();
    scene.add(character);

    const mat = (
        color,
        metalness = 0.35,
        roughness = 0.3,
        emissive = 0x000000
    ) => new THREE.MeshStandardMaterial({
        color,
        metalness,
        roughness,
        emissive,
        emissiveIntensity:
            visualStyle === "hologram" ? 0.45 : 0.12
    });

    const skinColor =
        visualStyle === "human" ? 0xe9ad89 :
        visualStyle === "cartoon" ? 0xffc4dc :
        visualStyle === "fantasy" ? 0xb9a0ff :
        visualStyle === "robot" ? 0x9eb7c9 :
        0x55dff4;

    const skin = mat(skinColor, 0.45, 0.28, accent);
    const dark = mat(0x101b34, 0.7, 0.25);

    const glow = new THREE.MeshBasicMaterial({
        color: accent
    });

    const white = new THREE.MeshStandardMaterial({
        color: 0xffffff,
        roughness: 0.2
    });

    const black = new THREE.MeshBasicMaterial({
        color: 0x07101d
    });

    function sphere(parent, material, position, scale) {

        const mesh = new THREE.Mesh(
            new THREE.SphereGeometry(1, 32, 24),
            material
        );

        mesh.position.set(...position);
        mesh.scale.set(...scale);

        parent.add(mesh);

        return mesh;
    }

    // TANA VA YELKA

    sphere(
        character,
        dark,
        [0, -1.45, 0],
        [0.92, 1.05, 0.48]
    );

    sphere(
        character,
        skin,
        [0, -0.78, 0.02],
        [0.86, 0.42, 0.48]
    );

    sphere(
        character,
        dark,
        [-0.82, -0.95, 0],
        [0.28, 0.72, 0.32]
    );

    sphere(
        character,
        dark,
        [0.82, -0.95, 0],
        [0.28, 0.72, 0.32]
    );

    // BO'YIN VA BOSH

    sphere(
        character,
        skin,
        [0, -0.28, 0],
        [0.28, 0.38, 0.3]
    );

    const head = new THREE.Group();

    head.position.set(0, 0.62, 0);
    character.add(head);

    sphere(
        head,
        skin,
        [0, 0, 0],
        [0.72, 0.88, 0.58]
    );

    // SOCH YOKI ROBOT QOPQOG'I

    if (
        visualStyle === "human" ||
        visualStyle === "cartoon" ||
        visualStyle === "fantasy"
    ) {

        sphere(
            head,
            dark,
            [0, 0.48, -0.02],
            [0.73, 0.38, 0.58]
        );

    } else {

        sphere(
            head,
            dark,
            [0, 0.55, 0],
            [0.76, 0.25, 0.61]
        );

        sphere(
            head,
            glow,
            [0, 0.77, 0],
            [0.18, 0.07, 0.12]
        );
    }

    // QULOQLAR

    sphere(
        head,
        dark,
        [-0.72, -0.02, 0],
        [0.13, 0.25, 0.16]
    );

    sphere(
        head,
        dark,
        [0.72, -0.02, 0],
        [0.13, 0.25, 0.16]
    );

    // KO'ZLAR

    const leftEye = sphere(
        head,
        white,
        [-0.27, 0.08, 0.49],
        [0.16, 0.19, 0.075]
    );

    const rightEye = sphere(
        head,
        white,
        [0.27, 0.08, 0.49],
        [0.16, 0.19, 0.075]
    );

    sphere(
        head,
        black,
        [-0.25, 0.07, 0.555],
        [0.075, 0.10, 0.035]
    );

    sphere(
        head,
        black,
        [0.29, 0.07, 0.555],
        [0.075, 0.10, 0.035]
    );

    // BURUN VA OG'IZ

    sphere(
        head,
        skin,
        [0, -0.12, 0.55],
        [0.075, 0.12, 0.08]
    );

    const mouth = sphere(
        head,
        black,
        [0, -0.34, 0.535],
        [0.19, 0.035, 0.035]
    );

    // KO'KRAK MARKAZI

    sphere(
        character,
        glow,
        [0, -1.25, 0.48],
        [0.16, 0.16, 0.06]
    );

    // AYLANUVCHI HALQALAR

    const ringMaterial = new THREE.MeshBasicMaterial({
        color: accent,
        transparent: true,
        opacity: 0.62
    });

    const ring1 = new THREE.Mesh(
        new THREE.TorusGeometry(
            1.42,
            0.012,
            8,
            100
        ),
        ringMaterial
    );

    ring1.rotation.x = 1.12;
    ring1.position.y = -0.55;

    character.add(ring1);

    const ring2 = new THREE.Mesh(
        new THREE.TorusGeometry(
            1.65,
            0.008,
            8,
            100
        ),
        ringMaterial
    );

    ring2.rotation.x = 1.35;
    ring2.rotation.y = 0.55;
    ring2.position.y = -0.55;

    character.add(ring2);

    // YULDUZCHALAR

    const starGeometry = new THREE.BufferGeometry();
    const starPositions = [];

    for (let i = 0; i < 180; i++) {

        starPositions.push(
            (Math.random() - 0.5) * 9,
            (Math.random() - 0.5) * 6,
            (Math.random() - 0.5) * 4 - 1
        );
    }

    starGeometry.setAttribute(
        'position',
        new THREE.Float32BufferAttribute(
            starPositions,
            3
        )
    );

    const stars = new THREE.Points(
        starGeometry,
        new THREE.PointsMaterial({
            color: accent,
            size: 0.025,
            transparent: true,
            opacity: 0.8
        })
    );

    scene.add(stars);

    // ANIMATSIYA

    let blinkAt = 0;
    let blinkUntil = 0;

    const clock = new THREE.Clock();

    function resize() {

        const width = stage.clientWidth;
        const height = stage.clientHeight;

        camera.aspect = width / height;
        camera.updateProjectionMatrix();

        renderer.setSize(width, height);
    }

    window.addEventListener('resize', resize);

    function animate() {

        requestAnimationFrame(animate);

        const t = clock.getElapsedTime();

        // Tana harakati
        character.position.y =
            Math.sin(t * 1.35) * 0.07;

        character.rotation.y =
            Math.sin(t * 0.55) * 0.13;

        // Bosh harakati
        head.rotation.y =
            Math.sin(t * 0.8) * 0.08;

        head.rotation.x =
            Math.sin(t * 0.65) * 0.035;

        // Halqalar
        ring1.rotation.z += 0.003;
        ring2.rotation.z -= 0.002;

        // Yulduzlar
        stars.rotation.y += 0.0005;

        // Ko'z pirpiratish
        if (t > blinkAt) {

            blinkUntil = t + 0.14;
            blinkAt = t + 3.2 + Math.random() * 2.2;
        }

        const blinking = t < blinkUntil;

        leftEye.scale.y = blinking ? 0.025 : 0.19;
        rightEye.scale.y = blinking ? 0.025 : 0.19;

        // Gapirish animatsiyasi
        if (isSpeaking) {

            mouth.scale.y =
                0.035 + Math.abs(Math.sin(t * 12)) * 0.12;

            mouth.scale.x =
                0.19 + Math.abs(Math.sin(t * 12)) * 0.06;

            statusNode.textContent = "JON AI GAPIRMOQDA";

        } else {

            mouth.scale.y = 0.035;
            mouth.scale.x = 0.19;

            statusNode.textContent = "VIRTUAL INTELLIGENCE";
        }

        renderer.render(scene, camera);
    }

    animate();

})();
</script>
</body>
</html>
"""

    html = (
        html.replace("__NAME__", safe_name)
            .replace("__STYLE__", visual_style)
            .replace(
                "__SPEAKING__",
                "true" if speaking else "false"
            )
    )

    components.html(
        html,
        height=height,
        scrolling=False
    )


# -------------------- POWERPOINT HELPERS --------------------

def parse_slides(result):

    try:

        clean = re.sub(
            r"```(?:json)?",
            "",
            result.strip(),
            flags=re.IGNORECASE
        )

        clean = clean.replace("```", "").strip()

        start = clean.find("[")
        end = clean.rfind("]") + 1

        if start < 0 or end <= start:
            raise ValueError("JSON massiv topilmadi")

        data = json.loads(clean[start:end])

        if not isinstance(data, list):
            raise ValueError("AI javobi ro'yxat shaklida emas")

        return data

    except Exception as exc:

        st.error(
            f"Slayd ma'lumotlarini o'qishda xatolik: {exc}"
        )

        return []


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
        body.clear()

        points = item.get("points", [])

        if isinstance(points, str):
            points = [points]

        for index, point in enumerate(points):

            paragraph = (
                body.paragraphs[0]
                if index == 0
                else body.add_paragraph()
            )

            paragraph.text = str(point)
            paragraph.level = 0

    output = BytesIO()

    prs.save(output)
    output.seek(0)

    return output


# -------------------- SIDEBAR --------------------

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

st.sidebar.title("💎 JON AI LIVE")
st.sidebar.caption("Virtual Intelligence Platform")

page = st.sidebar.radio(
    "Bo'limni tanlang",
    PAGES
)

st.sidebar.divider()

st.sidebar.caption(
    f"JON AI VERSION {APP_VERSION}"
)


# -------------------- HOME --------------------

if page == "🏠 Bosh sahifa":

    st.title("💎 JON AI LIVE")

    st.caption(
        "Sizning shaxsiy virtual yordamchingiz"
    )

    show_character()

    st.subheader("Xush kelibsiz!")

    st.write(
        "JON AI bilan suhbatlashing, bilim oling, "
        "tarjima qiling, prezentatsiya yarating "
        "va ma'lumotlarni tahlil qiling."
    )

    st.success(
        "Platforma ishga tayyor. "
        "AI ishlashi uchun Gemini API kaliti kerak."
    )


# -------------------- AI CHAT --------------------

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

            with st.spinner(
                "JON AI javob tayyorlamoqda..."
            ):

                answer = ask_ai(
                    prompt,
                    history=st.session_state.messages[:-1],
                )

            st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })

        st.rerun()


# -------------------- EDUCATION --------------------

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
            "Boshqa"
        ]
    )

    question = st.text_area(
        "Savolingiz yoki mavzuni yozing"
    )

    if st.button(
        "📚 Tushuntirish",
        use_container_width=True
    ):

        if question.strip():

            with st.spinner(
                "Javob tayyorlanmoqda..."
            ):

                answer = ask_ai(
                    question,
                    extra_context=(
                        f"Fan: {subject}. "
                        "Mavzuni tushunarli misollar bilan izohla."
                    ),
                )

            st.markdown(answer)

        else:

            st.warning(
                "Avval savolingizni yozing."
            )


# -------------------- POWERPOINT --------------------

elif page == "📊 PowerPoint":

    st.title("📊 JON AI — PowerPoint yaratish")

    topic = st.text_input(
        "Prezentatsiya mavzusi"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        3,
        15,
        7
    )

    if st.button(
        "✨ Prezentatsiya yaratish",
        use_container_width=True
    ):

        if topic.strip():

            prompt = f"""
            "{topic}" mavzusida {slide_count} ta slayd
            uchun JSON massiv yarat.

            Har bir elementda title va points maydonlari bo'lsin.
            points 3-5 ta qisqa banddan iborat bo'lsin.
            Faqat JSON qaytar.
            """

            with st.spinner(
                "Slaydlar tayyorlanmoqda..."
            ):

                

                                    result = ask_ai(prompt)

                slides = parse_slides(result)

                if slides:
                    pptx_file = create_presentation(
                        topic,
                        slides
                    )

                    st.download_button(
                        label="📥 PowerPoint yuklab olish",
                        data=pptx_file,
                        file_name="JON_AI_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                else:
                    st.error("Slaydlar yaratilmadi. Qayta urinib ko‘ring.")
          
