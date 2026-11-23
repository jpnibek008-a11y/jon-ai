
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

def show_character(height=520):
    config = {
        "height": height,
        "name": st.session_state.get("character_name", "JON"),
        "style": st.session_state.get("character_style", "Hologramma"),
        "mode": st.session_state.get("character_animation", "Tabiiy harakat"),
        "mood": st.session_state.get("jon_mood", "neutral"),
        "action": st.session_state.get("jon_action", "speaking"),
    }

    html_code = r"""
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
* {margin:0;padding:0;box-sizing:border-box}
body {
    overflow:hidden;
    background:transparent;
    font-family:Arial,sans-serif;
}
#scene {
    width:100%;
    height:CONFIG_HEIGHTpx;
    position:relative;
    overflow:hidden;
    border-radius:22px;
    background:
      radial-gradient(ellipse at 50% 55%,#112d59 0%,#09152d 42%,#050817 100%);
}
canvas {display:block}
#hud {
    position:absolute;
    top:16px;
    left:0;
    width:100%;
    text-align:center;
    pointer-events:none;
    z-index:3;
}
#name {
    color:#83eaff;
    font-size:24px;
    font-weight:800;
    letter-spacing:7px;
    text-shadow:0 0 15px #27caff;
}
#status {
    color:#7fffd4;
    font-size:10px;
    letter-spacing:3px;
    margin-top:7px;
}
#bottom {
    position:absolute;
    bottom:17px;
    left:0;
    width:100%;
    text-align:center;
    color:#8fc8ed;
    font-size:10px;
    letter-spacing:2px;
    pointer-events:none;
}
#glow {
    position:absolute;
    width:180px;
    height:180px;
    border-radius:50%;
    background:#00bfff;
    filter:blur(100px);
    opacity:.15;
    left:calc(50% - 90px);
    top:45%;
    pointer-events:none;
}
</style>
</head>
<body>
<div id="scene">
    <div id="glow"></div>
    <div id="hud">
        <div id="name">CONFIG_NAME</div>
        <div id="status">● JON AI — ONLINE</div>
    </div>
    <div id="bottom">CONFIG_STYLE · CONFIG_MODE</div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
<script>
(function(){
    const root = document.getElementById("scene");
    const status = document.getElementById("status");

    if (!window.THREE) {
        status.textContent = "3D MODUL YUKLANMADI";
        return;
    }

    const scene = new THREE.Scene();

    const camera = new THREE.PerspectiveCamera(
        42,
        root.clientWidth / root.clientHeight,
        0.1,
        100
    );
    camera.position.set(0,1.4,7);

    const renderer = new THREE.WebGLRenderer({
        antialias:true,
        alpha:true
    });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio,2));
    renderer.setSize(root.clientWidth,root.clientHeight);
    renderer.setClearColor(0x000000,0);
    root.appendChild(renderer.domElement);

    const group = new THREE.Group();
    scene.add(group);

    const ambient = new THREE.AmbientLight(0x8acfff,1.3);
    scene.add(ambient);

    const blueLight = new THREE.PointLight(0x00bfff,3,15);
    blueLight.position.set(-3,3,4);
    scene.add(blueLight);

    const purpleLight = new THREE.PointLight(0x8b42ff,2,15);
    purpleLight.position.set(3,1,2);
    scene.add(purpleLight);

    const whiteLight = new THREE.PointLight(0xffffff,1.5,10);
    whiteLight.position.set(0,4,-2);
    scene.add(whiteLight);

    const hologram = CONFIG_STYLE === "Hologramma";
    const cartoon = CONFIG_STYLE === "Multfilm uslubi";
    const fantasy = CONFIG_STYLE === "Fantastik";

    const mainColor = fantasy ? 0xb34dff :
                      cartoon ? 0x42dfff :
                      0x39c9ff;

    const metalColor = hologram ? 0x36dfff :
                       fantasy ? 0x9b55ff :
                       0x8aaacb;

    const material = new THREE.MeshStandardMaterial({
        color:metalColor,
        metalness:0.72,
        roughness:0.28,
        emissive:mainColor,
        emissiveIntensity:0.35,
        transparent:hologram,
        opacity:hologram ? 0.82 : 1
    });

    const darkMaterial = new THREE.MeshStandardMaterial({
        color:0x101a35,
        metalness:0.65,
        roughness:0.3
    });

    const glowMaterial = new THREE.MeshBasicMaterial({
        color:mainColor,
        transparent:true,
        opacity:0.85
    });

    function sphere(x,y,z,sx,sy,sz,mat) {
        const mesh = new THREE.Mesh(
            new THREE.SphereGeometry(1,32,24),
            mat
        );
        mesh.position.set(x,y,z);
        mesh.scale.set(sx,sy,sz);
        group.add(mesh);
        return mesh;
    }

    function cylinder(x,y,z,rt,rb,h,mat) {
        const mesh = new THREE.Mesh(
            new THREE.CylinderGeometry(rt,rb,h,32),
            mat
        );
        mesh.position.set(x,y,z);
        group.add(mesh);
        return mesh;
    }

    // BODY
    const body = sphere(0,-0.65,0,0.72,0.95,0.42,material);
    const chest = sphere(0,-0.35,0.36,0.43,0.43,0.13,glowMaterial);

    // NECK
    cylinder(0,0.42,0,0.19,0.22,0.36,darkMaterial);

    // HEAD
    const headGroup = new THREE.Group();
    headGroup.position.set(0,1.05,0);
    group.add(headGroup);

    const head = new THREE.Mesh(
        new THREE.SphereGeometry(0.65,40,32),
        material
    );
    head.scale.set(0.82,1,0.78);
    headGroup.add(head);

    // FACE PANEL
    const face = new THREE.Mesh(
        new THREE.SphereGeometry(0.49,32,24),
        new THREE.MeshBasicMaterial({
            color:0x06172c,
            transparent:true,
            opacity:0.85
        })
    );
    face.scale.set(0.95,0.7,0.22);
    face.position.set(0,-0.04,0.43);
    headGroup.add(face);

    // EYES
    const eyeMat = new THREE.MeshBasicMaterial({color:0x65edff});
    const eyeLeft = sphere(-0.22,1.08,0.46,0.105,0.045,0.035,eyeMat);
    const eyeRight = sphere(0.22,1.08,0.46,0.105,0.045,0.035,eyeMat);

    group.remove(eyeLeft);
    group.remove(eyeRight);
    headGroup.add(eyeLeft);
    headGroup.add(eyeRight);
    eyeLeft.position.set(-0.22,0.08,0.47);
    eyeRight.position.set(0.22,0.08,0.47);

    // MOUTH
    const mouth = new THREE.Mesh(
        new THREE.BoxGeometry(0.20,0.025,0.025),
        new THREE.MeshBasicMaterial({color:0x48eaff})
    );
    mouth.position.set(0,-0.20,0.49);
    headGroup.add(mouth);

    // ARMS
    const arms = [];
    for (let side of [-1,1]) {
        const arm = sphere(side*0.82,-0.55,0,0.19,0.65,0.20,material);
        arm.rotation.z = side * -0.18;
        arms.push(arm);

        const hand = sphere(side*0.88,-1.13,0.04,0.20,0.20,0.20,glowMaterial);
    }

    // SHOULDERS
    sphere(-0.72,-0.12,0,0.28,0.28,0.28,darkMaterial);
    sphere(0.72,-0.12,0,0.28,0.28,0.28,darkMaterial);

    // ENERGY RINGS
    const rings = [];
    for(let i=0;i<3;i++){
        const ring = new THREE.Mesh(
            new THREE.TorusGeometry(1.2+i*0.22,0.012,8,100),
            new THREE.MeshBasicMaterial({
                color:i===1 ? 0xa65cff : 0x39dfff,
                transparent:true,
                opacity:0.65-i*0.12
            })
        );
        ring.rotation.x = Math.PI/2 + i*0.18;
        ring.position.y = -1.45 + i*0.13;
        group.add(ring);
        rings.push(ring);
    }

    // FLOATING PARTICLES
    const particleCount = 180;
    const particlePositions = new Float32Array(particleCount*3);

    for(let i=0;i<particleCount;i++){
        particlePositions[i*3] = (Math.random()-0.5)*5;
        particlePositions[i*3+1] = (Math.random()-0.5)*5;
        particlePositions[i*3+2] = (Math.random()-0.5)*3;
    }

    const particleGeometry = new THREE.BufferGeometry();
    particleGeometry.setAttribute(
        "position",
        new THREE.BufferAttribute(particlePositions,3)
    );

    const particles = new THREE.Points(
        particleGeometry,
        new THREE.PointsMaterial({
            color:0x58dfff,
            size:0.035,
            transparent:true,
            opacity:0.8
        })
    );
    scene.add(particles);

    // LIGHT BASE
    const base = new THREE.Mesh(
        new THREE.CylinderGeometry(1.1,1.3,0.12,64),
        new THREE.MeshBasicMaterial({
            color:0x164b82,
            transparent:true,
            opacity:0.7
        })
    );
    base.position.y = -1.72;
    group.add(base);

    const clock = new THREE.Clock();

    function animate(){
        requestAnimationFrame(animate);

        const t = clock.getElapsedTime();
        const mode = CONFIG_MODE;
        const mood = CONFIG_MOOD;
        const action = CONFIG_ACTION;

        // BODY MOVEMENT
        group.position.y = Math.sin(t*1.5)*0.07;

        if(mode === "Raqs"){
            group.rotation.y = Math.sin(t*1.8)*0.28;
            group.rotation.z = Math.sin(t*3)*0.08;
            arms[0].rotation.z = -0.3 + Math.sin(t*4)*0.35;
            arms[1].rotation.z = 0.3 + Math.cos(t*4)*0.35;
        } else if(mode === "Energetik"){
            group.rotation.y = Math.sin(t*2.5)*0.18;
            arms[0].rotation.z = -0.3 + Math.sin(t*3)*0.2;
            arms[1].rotation.z = 0.3 + Math.cos(t*3)*0.2;
        } else if(mode === "Sokin"){
            group.rotation.y = Math.sin(t*0.5)*0.06;
        } else {
            group.rotation.y = Math.sin(t*0.8)*0.13;
            arms[0].rotation.z = -0.18 + Math.sin(t*1.3)*0.08;
            arms[1].rotation.z = 0.18 + Math.cos(t*1.3)*0.08;
        }

        // HEAD EXPRESSION
        headGroup.rotation.y = Math.sin(t*0.8)*0.12;
        headGroup.rotation.x = Math.sin(t*0.7)*0.06;

        if(mood === "thinking"){
            headGroup.rotation.x = -0.15 + Math.sin(t)*0.04;
        }
        if(mood === "happy" || mood === "greeting"){
            headGroup.rotation.z = Math.sin(t*2)*0.04;
        } else {
            headGroup.rotation.z = Math.sin(t*0.7)*0.02;
        }

        // BLINKING
        const blink = (Math.sin(t*1.7)>0.985) ? 0.15 : 1;
        eyeLeft.scale.y = 0.045*blink;
        eyeRight.scale.y = 0.045*blink;

        // SPEAKING MOUTH
        if(action === "speaking" || mood === "greeting"){
            mouth.scale.y = 1 + Math.abs(Math.sin(t*10))*5;
        } else {
            mouth.scale.y = 1;
        }

        // RING MOVEMENT
        rings.forEach((ring,i)=>{
            ring.rotation.z = t*(0.25+i*0.12)*(i%2===0 ? 1 : -1);
        });

        particles.rotation.y = t*0.025;
        particles.rotation.x = Math.sin(t*0.3)*0.08;

        blueLight.intensity = 2.5 + Math.sin(t*2)*0.8;

        renderer.render(scene,camera);
    }

    animate();

    window.addEventListener("resize",()=>{
        const w = root.clientWidth;
        const h = root.clientHeight;
        camera.aspect = w/h;
        camera.updateProjectionMatrix();
        renderer.setSize(w,h);
    });

})();
</script>
</div>
</body>
</html>
"""

    replacements = {
        "CONFIG_HEIGHT": str(height),
        "CONFIG_NAME": json.dumps(config["name"], ensure_ascii=False)[1:-1],
        "CONFIG_STYLE": json.dumps(config["style"], ensure_ascii=False)[1:-1],
        "CONFIG_MODE": json.dumps(config["mode"], ensure_ascii=False)[1:-1],
        "CONFIG_MOOD": json.dumps(config["mood"], ensure_ascii=False),
        "CONFIG_ACTION": json.dumps(config["action"], ensure_ascii=False),
    }

    for key, value in replacements.items():
        html_code = html_code.replace(key, value)

    components.html(html_code, height=height + 10, scrolling=False)

# ---------------------- NAVIGATION ----------------------

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

    DEFAULT_MODEL = "gemini-3.5-flash-lite"

    model_options = [
        "gemini-3.5-flash-lite",
    ]

    current_model = st.session_state.get("model", DEFAULT_MODEL)

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
    
        
