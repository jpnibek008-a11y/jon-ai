
# ============================================================
# JON AI LIVE 9.2 — NEXT GENERATION VIRTUAL INTELLIGENCE
# Streamlit + Gemini AI + Three.js Dual 3D Characters
# ============================================================

import json
import re
import time
import random
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt


# ============================================================
# 1. CONFIGURATION
# ============================================================

APP_VERSION = "9.3"
DEFAULT_MODEL = "gemini-3.8-flash"

st.set_page_config(
    page_title="JON AI LIVE 9.3",
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
Foydalanuvchiga asosan o'zbek tilida javob bering.
Ingliz, rus va boshqa tillarda ham yordam bera olasiz.
Bilmagan ma'lumotni to'qib chiqarmang.
Ta'lim, dasturlash, biznes, tarjima, ijod va tahlilda yordam bering.
Javoblaringiz aniq, tartibli va amaliy bo'lsin.
Murakkab vazifalarni bosqichlarga ajrating.
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
    "character_animation": "Raqs",
    "copied_prompt": "",
    "last_answer": "",
    "motivation_quote": ""
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value

MOTIVATIONAL_QUOTES = [
    ("TO‘XTAMA. SEN HALI ORZU QILGAN MANZILINGGA YETMAGANSAN!",
     "Charchashing mumkin. Dam olishing mumkin. Ammo taslim bo‘lma."),
    ("BUGUNGI HARAKATING — ERTANGI G‘ALABANG.",
     "Kichik qadamlar ham seni katta maqsadga yaqinlashtiradi."),
    ("QO‘RQMA. BOSHLAGAN ISHINGNI OXIRIGA YETKAZ!",
     "Mukammal bo‘lish shart emas. Muhimi — davom etish."),
    ("SEN O‘YLAGANINGDAN HAM KUCHLISAN!",
     "Qiyinchiliklar seni to‘xtatish uchun emas, chiniqtirish uchun."),
    ("BUGUN TASLIM BO‘LMA — ERTANGI O‘ZING SENGA RAHMAT AYTADI.",
     "Maqsadingni esla va yana bir qadam tashla."),
    ("ORZULAR KUTMAYDI. HARAKATNI BUGUN BOSHLА!",
     "Vaqtingni qadrlagin. O‘z tarixingni o‘zing yarat.")
]

if not st.session_state.get("motivation_quote"):
    st.session_state.motivation_quote = random.choice(MOTIVATIONAL_QUOTES)


# ============================================================
# 3. DESIGN
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@500;700;800&family=Inter:wght@300;400;500;600;700&display=swap');

:root {
    --blue:#26baff;
    --purple:#9b5cff;
    --bg:#050817;
}

.stApp {
    background:
    radial-gradient(ellipse at 15% 10%,#17265a 0%,transparent 35%),
    radial-gradient(ellipse at 85% 20%,#25124d 0%,transparent 35%),
    linear-gradient(145deg,#050817,#080d20 60%,#09051c);
    color:#edf4ff;
    font-family:Inter,sans-serif;
}

h1,h2,h3 {color:#edf4ff !important;}

.hero-title {
    font-family:Orbitron,sans-serif;
    font-weight:800;
    letter-spacing:4px;
    text-align:center;
    font-size:clamp(28px,5vw,58px);
    background:linear-gradient(90deg,#fff,#43caff,#a879ff);
    -webkit-background-clip:text;
    -webkit-text-fill-color:transparent;
}

.hero-subtitle {
    color:#9bb8df;
    text-align:center;
    letter-spacing:4px;
    font-size:11px;
}

.panel {
    background:linear-gradient(145deg,#111c3bdb,#10102bdd);
    border:1px solid #4e8cff44;
    border-radius:18px;
    padding:20px;
    margin-bottom:15px;
    box-shadow:0 0 25px #176dff12;
}

.motivation-frame {
    position:relative;
    overflow:hidden;
    margin:22px auto 24px;
    padding:20px 24px;
    max-width:980px;
    text-align:center;
    border:1px solid rgba(72,205,255,.72);
    border-radius:20px;
    background:linear-gradient(125deg,rgba(9,28,58,.94),rgba(32,17,65,.92));
    box-shadow:0 0 14px rgba(38,186,255,.22),inset 0 0 24px rgba(80,155,255,.10);
    animation:motivationGlow 4s ease-in-out infinite;
}
.motivation-frame:before {
    content:"";
    position:absolute;
    inset:5px;
    border:1px solid rgba(157,111,255,.34);
    border-radius:15px;
    pointer-events:none;
}
.motivation-label {
    color:#65dcff;
    font-size:11px;
    font-weight:800;
    letter-spacing:3px;
    margin-bottom:10px;
}
.motivation-quote {
    position:relative;
    color:#fff;
    font-family:Orbitron,Inter,sans-serif;
    font-size:clamp(15px,2.2vw,23px);
    line-height:1.55;
    font-weight:800;
    text-shadow:0 0 14px rgba(65,194,255,.34);
}
.motivation-subtext {
    position:relative;
    color:#b9cce8;
    font-size:13px;
    margin-top:9px;
}
.motivation-signature {
    position:relative;
    color:#72dfff;
    font-size:11px;
    font-weight:700;
    letter-spacing:3px;
    margin-top:13px;
}
@keyframes motivationGlow {
    0%,100% {box-shadow:0 0 12px rgba(38,186,255,.18),inset 0 0 20px rgba(80,155,255,.08)}
    50% {box-shadow:0 0 25px rgba(38,186,255,.34),inset 0 0 28px rgba(130,90,255,.13)}
}
@media(max-width:600px) {
    .motivation-frame {padding:17px 14px;margin:16px auto 19px;border-radius:16px}
    .motivation-label {font-size:10px;letter-spacing:2px}
    .motivation-subtext {font-size:12px}
}

.feature-card {
    background:linear-gradient(145deg,#111c38,#15102e);
    border:1px solid #438dff44;
    border-radius:15px;
    padding:18px;
    min-height:115px;
}

div.stButton > button {
    border-radius:12px;
    border:1px solid #438dff77;
    background:linear-gradient(100deg,#123c75,#39206b);
    color:white;
    font-weight:600;
    min-height:42px;
}

div.stButton > button:hover {
    border-color:#54d5ff;
    color:white;
    box-shadow:0 0 15px #36baff55;
}

[data-testid="stSidebar"] {
    background:linear-gradient(180deg,#080e23,#100d29);
    border-right:1px solid #3475b944;
}

[data-testid="stChatMessage"] {
    background:#101a35;
    border:1px solid #3d6eb533;
    border-radius:14px;
}

.stTextInput input,.stTextArea textarea {
    background:#0d1730 !important;
    color:white !important;
    border-color:#3475b966 !important;
}

hr {border-color:#3c69a633;}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 4. GEMINI AI ENGINE
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
            "API so'rov limiti yoki kvotasi tugagan. "
            "Google AI Studio'dagi limitlarni tekshiring."
        )

    if "503" in message or "unavailable" in lower:
        return "Gemini serveri vaqtincha band. Keyinroq qayta urinib ko'ring."

    if "401" in message or "403" in message or "api key" in lower:
        return "API kalit noto'g'ri yoki uning ruxsatlari yetarli emas."

    if "404" in message or "not found" in lower:
        return (
            "Tanlangan model topilmadi. Sozlamalardan "
            "gemini-2.5-flash modelini tanlab ko'ring."
        )

    if "timeout" in lower:
        return "Ulanish vaqti tugadi. Internetni tekshiring."

    return f"Gemini xatosi: {message[:400]}"


def ask_ai(prompt, extra_context="", history=None):
    try:
        client = get_client()
        model = st.session_state.get("model", DEFAULT_MODEL)

        conversation = [JON_CORE]

        for item in (history or [])[-12:]:
            role = item.get("role", "user")
            content = item.get("content", "")

            if content:
                role_name = (
                    "Foydalanuvchi"
                    if role == "user"
                    else "JON AI"
                )
                conversation.append(f"{role_name}: {content}")

        if extra_context:
            conversation.append(
                f"Qo'shimcha kontekst:\n{extra_context}"
            )

        conversation.append(f"Foydalanuvchi: {prompt}")
        conversation.append("JON AI:")

        response = client.models.generate_content(
            model=model,
            contents="\n\n".join(conversation)
        )

        answer = getattr(response, "text", None)

        if not answer:
            return "AI javob qaytarmadi. Qayta urinib ko'ring."

        return answer.strip()

    except Exception as error:
        return explain_api_error(error)


def check_model():
    try:
        client = get_client()
        model = st.session_state.get("model", DEFAULT_MODEL)

        response = client.models.generate_content(
            model=model,
            contents="Salom. Faqat OK deb javob ber."
        )

        return bool(getattr(response, "text", None))

    except Exception as error:
        st.error(explain_api_error(error))
        return False

            return "AI javob qaytarmadi. Qayta urinib ko'ring."

        return answer.strip()

    except Exception as error:
        return explain_api_error(error)


def check_model():
    try:
        client = get_client()
        model = st.session_state.get("model", DEFAULT_MODEL)

        response = client.models.generate_content(
            model=model,
            contents="Salom. Faqat OK deb javob ber."
        )

        return bool(getattr(response, "text", None))

    except Exception as error:
        st.error(explain_api_error(error))
        return False


# ============================================================
# 5. THREE.JS — HUMAN + ANDROID
# ============================================================

def show_character(height=610, cinematic=True):

    name = json.dumps(
        st.session_state.get("character_name", "JON")
    )

    style = json.dumps(
        st.session_state.get("character_style", "Hologramma")
    )

    animation = json.dumps(
        st.session_state.get("character_animation", "Raqs")
    )

    intro = (
        "true"
        if cinematic and st.session_state.get("intro_enabled", True)
        else "false"
    )

    html = r"""
<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
*{box-sizing:border-box}
html,body{
 margin:0;width:100%;height:100%;
 overflow:hidden;background:transparent;
 font-family:Arial,sans-serif
}
#scene{
 width:100%;height:100%;position:relative;
 overflow:hidden;border-radius:18px;
 background:
 radial-gradient(ellipse at 50% 60%,#142b58aa,transparent 55%),
 radial-gradient(ellipse at 50% 20%,#21134b99,transparent 65%),
 #050817;
}
canvas{display:block;width:100%;height:100%}
.brand{
 position:absolute;top:15px;width:100%;
 text-align:center;z-index:3;pointer-events:none
}
.brand h1{
 margin:0;color:white;font-size:32px;
 letter-spacing:7px;text-shadow:0 0 20px #36caff
}
.brand span{
 color:#a8c8ef;font-size:9px;letter-spacing:4px
}
.status{
 position:absolute;top:78px;left:50%;
 transform:translateX(-50%);z-index:3;
 color:#5effce;font-size:10px;letter-spacing:3px;
}
.labels{
 position:absolute;bottom:17px;width:100%;
 display:flex;justify-content:space-around;
 z-index:3;pointer-events:none
}
.label{
 text-align:center;color:#c5e8ff;
 font-size:11px;letter-spacing:3px
}
.label small{
 display:block;margin-top:5px;color:#7b9dc9;
 font-size:8px;letter-spacing:2px
}
#intro{
 position:absolute;inset:0;z-index:5;
 display:flex;align-items:center;justify-content:center;
 background:radial-gradient(circle,#173b75,#090d22 65%,#02040c);
 color:white;font-size:14px;letter-spacing:5px;
 transition:opacity 1.5s;pointer-events:none
}
#intro.hide{opacity:0}
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
(function(){
 const box=document.getElementById("scene");
 const intro=document.getElementById("intro");
 const status=document.getElementById("status");

 const characterName=__NAME__;
 const characterStyle=__STYLE__;
 const animationMode=__ANIMATION__;
 const introEnabled=__INTRO__;

 const greetingStatus=document.getElementById("status");
 greetingStatus.textContent="● SALOM, "+characterName.toUpperCase()+"!";
 setTimeout(()=>{
  if(greetingStatus) greetingStatus.textContent="● SYSTEM ONLINE";
 },3500);

 if(!window.THREE){
  intro.innerText="3D ENGINE LOAD FAILED";
  return;
 }

 const THREE=window.THREE;
 const scene=new THREE.Scene();
 scene.fog=new THREE.FogExp2(0x050817,0.025);

 const camera=new THREE.PerspectiveCamera(
  42,box.clientWidth/box.clientHeight,0.1,100
 );
 camera.position.set(0,1.0,10.8);

 const renderer=new THREE.WebGLRenderer({
  antialias:true,alpha:true
 });
 renderer.setPixelRatio(Math.min(devicePixelRatio||1,1.5));
 renderer.setSize(box.clientWidth,box.clientHeight);
 renderer.outputEncoding=THREE.sRGBEncoding;
 box.appendChild(renderer.domElement);

 scene.add(new THREE.AmbientLight(0x7899ff,1.5));

 const light=new THREE.PointLight(0x38bfff,3,30);
 light.position.set(-4,5,5);
 scene.add(light);

 const purpleLight=new THREE.PointLight(0x974bff,2.5,25);
 purpleLight.position.set(4,3,3);
 scene.add(purpleLight);

 const floorLight=new THREE.PointLight(0x176bff,2,20);
 floorLight.position.set(0,-3,4);
 scene.add(floorLight);

 let humanColor=0x167bd1;
 let androidColor=0x101d38;

 if(characterStyle==="Realistik robot"){
  humanColor=0x9b8d82;
  androidColor=0x52677d;
 }
 if(characterStyle==="Multfilm uslubi"){
  humanColor=0x40aaff;
  androidColor=0x6e45ff;
 }
 if(characterStyle==="Fantastik"){
  humanColor=0x8c48e8;
  androidColor=0x00bfff;
 }

 function mat(color,metal=0.5,rough=0.3){
  return new THREE.MeshStandardMaterial({
   color,metalness:metal,roughness:rough,
   emissive:color,emissiveIntensity:0.16
  });
 }

 const blue=mat(humanColor,0.7,0.25);
 const cyan=mat(0x55dfff,0.6,0.2);
 const dark=mat(androidColor,0.8,0.28);
 const skin=mat(0xc58e78,0.1,0.55);
 const purple=mat(0x8c48e8,0.7,0.2);

 function sphere(parent,x,y,z,sx,sy,sz,m){
  const mesh=new THREE.Mesh(
   new THREE.SphereGeometry(1,24,18),m
  );
  mesh.position.set(x,y,z);
  mesh.scale.set(sx,sy,sz);
  parent.add(mesh);
  return mesh;
 }

 function cylinder(parent,x,y,z,rt,rb,h,m){
  const mesh=new THREE.Mesh(
   new THREE.CylinderGeometry(rt,rb,h,20),m
  );
  mesh.position.set(x,y,z);
  parent.add(mesh);
  return mesh;
 }

 function ring(parent,r,t,color,x,y,z,rx){
  const mesh=new THREE.Mesh(
   new THREE.TorusGeometry(r,t,10,90),
   new THREE.MeshBasicMaterial({
    color,transparent:true,opacity:0.75
   })
  );
  mesh.position.set(x,y,z);
  mesh.rotation.x=rx;
  parent.add(mesh);
  return mesh;
 }

 function createPerson(isAndroid){
  const root=new THREE.Group();
  root.position.x=isAndroid?2.0:-2.0;
  scene.add(root);

  const body=new THREE.Group();
  root.add(body);

  const bodyMat=isAndroid?dark:blue;
  const faceMat=isAndroid?dark:skin;

  cylinder(body,-0.22,-1.05,0,0.14,0.18,0.9,bodyMat);
  cylinder(body,0.22,-1.05,0,0.14,0.18,0.9,bodyMat);

  sphere(body,-0.22,-1.48,0.1,0.2,0.12,0.3,blue);
  sphere(body,0.22,-1.48,0.1,0.2,0.12,0.3,blue);

  sphere(body,0,-0.05,0,0.57,0.78,0.32,bodyMat);
  sphere(body,0,-0.08,0.3,0.35,0.45,0.1,blue);

  cylinder(body,0,0.72,0,0.15,0.17,0.32,bodyMat);

  const head=new THREE.Group();
  head.position.y=1.35;
  body.add(head);

  sphere(
   head,0,0,0,
   isAndroid?0.45:0.43,
   isAndroid?0.48:0.55,
   0.36,faceMat
  );

  if(isAndroid){
   sphere(head,0,0.02,0.29,0.36,0.25,0.12,blue);
   sphere(head,0,0.02,0.39,0.29,0.13,0.035,cyan);
   sphere(head,-0.43,0,0,0.07,0.15,0.13,cyan);
   sphere(head,0.43,0,0,0.07,0.15,0.13,cyan);
  }else{
   sphere(head,0,0.39,-0.015,0.39,0.19,0.34,dark);
  }

  const eyeMat=new THREE.MeshBasicMaterial({color:0xb9f8ff});
  const eyeL=sphere(head,-0.15,0.07,0.34,0.06,0.035,0.025,eyeMat);
  const eyeR=sphere(head,0.15,0.07,0.34,0.06,0.035,0.025,eyeMat);
  const mouth=sphere(head,0,-0.19,0.36,0.1,0.018,0.02,purple);

  const armL=new THREE.Group();
  const armR=new THREE.Group();

  armL.position.set(-0.57,0.36,0);
  armR.position.set(0.57,0.36,0);
  body.add(armL,armR);

  const armMat=isAndroid?blue:skin;

  cylinder(armL,0,-0.38,0,0.14,0.17,0.7,armMat);
  cylinder(armR,0,-0.38,0,0.14,0.17,0.7,armMat);

  sphere(armL,0,-0.78,0,0.15,0.16,0.15,armMat);
  sphere(armR,0,-0.78,0,0.15,0.16,0.15,armMat);

  const rings=[
   ring(root,0.9,0.018,0x45cfff,0,-1.7,0,0),
   ring(root,0.75,0.012,0x8e55ff,0,-1.7,0,0)
  ];

  return {root,body,head,mouth,eyeL,eyeR,armL,armR,rings};
 }

 const human=createPerson(false);
 const android=createPerson(true);

 const energyRings=[];

 for(let i=0;i<5;i++){
  energyRings.push(
   ring(scene,2.3+i*0.24,0.008,
    i%2?0x9655ff:0x36cfff,
    0,-0.1,-1.2,Math.PI/2)
  );
 }

 const portal=new THREE.Group();
 portal.position.set(0,0.1,-1.5);
 scene.add(portal);

 ring(portal,3.15,0.035,0x4d9fff,0,0,0,Math.PI/2);

 const count=innerWidth<600?250:500;
 const positions=new Float32Array(count*3);
    
        prompt = f"""
        "{title}" mavzusida {count} slaydli PowerPoint taqdimoti uchun
        mazmun tayyorla.

        Til: {language}
        Dizayn uslubi: {style}

        Natijani quyidagi formatda ber:
        SLIDE 1: Sarlavha
        - Birinchi fikr
        - Ikkinchi fikr

        Har bir slayd uchun sarlavha va 3-5 ta qisqa mazmun yoz.
        """

        with st.spinner("Taqdimot mazmuni tayyorlanmoqda..."):
            content = ask_ai(prompt)

        st.markdown(content)

        presentation = Presentation()
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)

        slides_data = re.split(
            r"(?=SLIDE\s+\d+\s*:)",
            content,
            flags=re.IGNORECASE
        )

        for index, slide_text in enumerate(slides_data):
            if not slide_text.strip():
                continue

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[6]
            )

            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(8, 17, 40)

            lines = [
                line.strip()
                for line in slide_text.splitlines()
                if line.strip()
            ]

            heading = lines[0] if lines else f"Slayd {index + 1}"

            title_box = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.5),
                Inches(12), Inches(1.0)
            )

            title_frame = title_box.text_frame
            title_frame.text = heading

            title_run = title_frame.paragraphs[0].runs[0]
            title_run.font.size = Pt(30)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(70, 205, 255)

            body_box = slide.shapes.add_textbox(
                Inches(1.0), Inches(2.0),
                Inches(11.3), Inches(4.5)
            )

            body_frame = body_box.text_frame
            body_frame.word_wrap = True

            for line in lines[1:]:
                paragraph = body_frame.add_paragraph()
                paragraph.text = line
                paragraph.font.size = Pt(21)
                paragraph.font.color.rgb = RGBColor(235, 243, 255)
                paragraph.space_after = Pt(14)

        output = BytesIO()
        presentation.save(output)
        output.seek(0)

        st.download_button(
            "📥 PowerPoint faylini yuklab olish",
            data=output,
            file_name="JON_AI_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )


# ============================================================
# 11. TRANSLATOR
# ============================================================

def render_translator():
    st.title("🌐 JON AI — Professional Translator")

    mode = st.radio(
        "Tarjima turi",
        ["Zamonaviy tillar", "Qadimgi tillar"],
        horizontal=True
    )

    if mode == "Zamonaviy tillar":
        languages = [
            "O'zbek", "English", "Русский", "Türkçe",
            "العربية", "中文", "日本語", "한국어",
            "Français", "Deutsch", "Español", "Italiano"
        ]
    else:
        languages = [
            "Classical Latin",
            "Ancient Greek",
            "Sanskrit",
            "Biblical Hebrew",
            "Aramaic",
            "Ancient Egyptian hieroglyphs",
            "Akkadian",
            "Sumerian",
            "Old Turkic runic"
        ]

    source_language = st.selectbox(
        "Qaysi tildan?",
        ["Avtomatik aniqlash"] + languages
    )

    target_language = st.selectbox(
        "Qaysi tilga?",
        languages
    )

    text_input = st.text_area(
        "Tarjima qilinadigan matn",
        height=180
    )

    if mode == "Qadimgi tillar":
        st.info(
            "Qadimgi tillar tarjimasida ayrim so'zlarning ma'nosi "
            "taxminiy bo'lishi mumkin. JON AI buni alohida ko'rsatadi."
        )

    if st.button("🌍 Tarjima qilish", use_container_width=True):
        if not text_input.strip():
            st.warning("Tarjima uchun matn kiriting.")
            return

        if mode == "Qadimgi tillar":
            prompt = f"""
            Sen qadimgi tillar bo'yicha ehtiyotkor tarjimon bo'l.

            Manba tili: {source_language}
            Maqsad tili: {target_language}

            Matn:
            {text_input}

            Quyidagi tartibda javob ber:
            1. Tarjima.
            2. So'zma-so'z tahlil.
            3. Talaffuz yoki transliteratsiya.
            4. Tarixiy va grammatik izoh.
            5. Qaysi joylari aniq, qaysilari taxminiy ekanini ko'rsat.

            Mavjud bo'lmagan qadimgi so'z yoki yozuvni o'ylab topma.
            Ishonchli ma'lumot bo'lmasa, buni ochiq ayt.
            """
        else:
            prompt = f"""
            Matnni {source_language} tilidan {target_language} tiliga
            tabiiy va mazmunini saqlagan holda tarjima qil.

            Matn:
            {text_input}
            """

        with st.spinner("Tarjima qilinmoqda..."):
            answer = ask_ai(prompt)

        st.markdown("### Natija")
        st.markdown(answer)

        st.session_state.last_answer = answer

        st.download_button(
            "📥 Tarjimani TXT shaklida yuklab olish",
            data=answer,
            file_name="JON_AI_translation.txt",
            mime="text/plain"
        )


# ============================================================
# 12. IMAGE PROMPT STUDIO
# ============================================================

def render_image_prompt():
    st.title("🎨 JON AI — Image Prompt Studio")

    subject = st.text_input(
        "Rasm mavzusi",
        placeholder="Masalan: kelajakdagi aqlli shahar"
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
            "Corporate"
        ]
    )

    ratio = st.selectbox(
        "Rasm o'lchami",
        ["1:1", "16:9", "9:16", "4:3", "3:2"]
    )

    details = st.text_area("Qo'shimcha tafsilotlar")

    if st.button("✨ Prompt yaratish", use_container_width=True):
        if not subject.strip():
            st.warning("Rasm mavzusini kiriting.")
            return

        prompt = f"""
        Professional image-generation prompt yoz.

        Mavzu: {subject}
        Uslub: {style}
        Aspect ratio: {ratio}
        Tafsilotlar: {details}

        Natijada:
        - Ingliz tilidagi tayyor prompt.
        - Yorug'lik va kamera tavsifi.
        - Ranglar va kompozitsiya.
        - Negative prompt.

        Prompt aniq, batafsil va professional bo'lsin.
        """

        with st.spinner("Prompt yaratilmoqda..."):
            answer = ask_ai(prompt)

        st.markdown(answer)

        st.session_state.copied_prompt = answer

        st.code(answer, language="text")

        st.button(
            "📋 Promptni nusxalash",
            on_click=st.session_state.update,
            args=({"copied_prompt": answer},)
        )


# ============================================================
# 13. EXCEL ANALYSIS
# ============================================================

def render_excel():
    st.title("📈 JON AI — Excel Data Analysis")

    uploaded = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "xls", "csv"]
    )

    if not uploaded:
        st.info("Tahlil qilish uchun fayl yuklang.")
        return

    try:
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded)

        st.success("Fayl muvaffaqiyatli o'qildi.")

        st.markdown("### Jadval")
        st.dataframe(df, use_container_width=True)

        c1, c2, c3 = st.columns(3)
        c1.metric("Qatorlar", len(df))
        c2.metric("Ustunlar", len(df.columns))
        c3.metric("Bo'sh qiymatlar", int(df.isna().sum().sum()))

        st.markdown("### Statistik ma'lumot")
        st.dataframe(
            df.describe(include="all").fillna(""),
            use_container_width=True
        )

        question = st.text_area(
            "JON AI ga qanday tahlil kerak?",
            placeholder="Jadvaldagi asosiy tendensiyalarni tushuntir..."
        )

        if st.button("🤖 AI bilan tahlil qilish"):
            sample = df.head(40).to_string()

            prompt = f"""
            Quyidagi jadvalni tahlil qil.

            Ustunlar: {list(df.columns)}
            Qatorlar soni: {len(df)}

            Namuna:
            {sample}

            Foydalanuvchi talabi:
            {question}

            Xulosalarni dalillarga asosla.
            Yetishmayotgan ma'lumotni o'ylab topma.
            """

            with st.spinner("Jadval tahlil qilinmoqda..."):
                answer = ask_ai(prompt)

            st.markdown(answer)

        csv_data = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "📥 CSV faylni yuklab olish",
            data=csv_data,
            file_name="JON_AI_data.csv",
            mime="text/csv"
        )

    except Exception as error:
        st.error(f"Faylni o'qishda xatolik: {error}")


# ============================================================
# 14. TASK MANAGER
# ============================================================

def render_tasks():
    st.title("📝 JON AI — Vazifalar boshqaruvi")

    task_text = st.text_input(
        "Yangi vazifa",
        placeholder="Masalan: JON AI kodini tekshirish"
    )

    priority = st.selectbox(
        "Muhimlik darajasi",
        ["Oddiy", "Muhim", "Juda muhim"]
    )

    if st.button("➕ Vazifa qo'shish"):
        if task_text.strip():
            st.session_state.tasks.append({
                "text": task_text.strip(),
                "priority": priority,
                "done": False
            })
            st.success("Vazifa qo'shildi.")
            st.rerun()

    st.markdown("---")
    st.markdown("### Mening vazifalarim")

    if not st.session_state.tasks:
        st.info("Hozircha vazifalar yo'q.")
        return

    for index, task in enumerate(st.session_state.tasks):
        col1, col2 = st.columns([5, 1])

        with col1:
            checked = st.checkbox(
                f"{task['text']} — {task['priority']}",
                value=task["done"],
                key=f"task_{index}"
            )
            st.session_state.tasks[index]["done"] = checked

        with col2:
            if st.button("🗑️", key=f"delete_task_{index}"):
                st.session_state.tasks.pop(index)
                st.rerun()

    completed = sum(
        1 for task in st.session_state.tasks if task["done"]
    )

    st.progress(completed / len(st.session_state.tasks))
    st.caption(
        f"{completed} / {len(st.session_state.tasks)} vazifa bajarildi."
    )
    
 const clock=new THREE.Clock();

 function animate(){
  requestAnimationFrame(animate);
  const t=clock.getElapsedTime();

  [human,android].forEach((p,i)=>{
   const phase=i*Math.PI;
   let speed=1.0, bounce=0.10, armSwing=0.22, bodySwing=0.025;

   if(animationMode==="Raqs"){
    speed=2.0; bounce=0.17; armSwing=0.62; bodySwing=0.10;
   }else if(animationMode==="Tabiiy harakat"){
    speed=1.15; bounce=0.09; armSwing=0.25; bodySwing=0.035;
   }else if(animationMode==="Sokin"){
    speed=0.65; bounce=0.035; armSwing=0.08; bodySwing=0.012;
   }else if(animationMode==="Energetik"){
    speed=2.8; bounce=0.22; armSwing=0.78; bodySwing=0.14;
   }

   p.root.position.y=Math.sin(t*speed+phase)*bounce;
   p.body.rotation.z=Math.sin(t*speed*0.8+phase)*bodySwing;
   p.head.rotation.y=Math.sin(t*0.8+phase)*0.12;
   p.head.rotation.x=Math.sin(t*1.5+phase)*0.035;

   p.armL.rotation.z=
    Math.sin(t*speed*1.7+phase)*armSwing+0.08;

   p.armR.rotation.z=
    -Math.sin(t*speed*1.7+phase)*armSwing-0.08;

   p.root.rotation.y=Math.sin(t*0.5+phase)*0.08;

   const blink=(Math.sin(t*2.4+phase)>0.985)?0.08:1;
   p.eyeL.scale.y=0.035*blink;
   p.eyeR.scale.y=0.035*blink;

   p.mouth.scale.y=0.018+Math.abs(Math.sin(t*3+phase))*0.025;

   p.rings.forEach((r,j)=>{
    r.rotation.z=t*(j?0.4:-0.3);
   });
  });

  energyRings.forEach((r,i)=>{
   r.rotation.z=t*(i%2?0.12:-0.12);
   r.material.opacity=0.4+Math.sin(t*2+i)*0.3;
  });

  portal.rotation.z=t*0.08;
  particles.rotation.y=t*0.025;
  particles.rotation.x=Math.sin(t*0.2)*0.08;

  light.intensity=2.4+Math.sin(t*2)*0.5;
  purpleLight.intensity=2+Math.cos(t*1.7)*0.5;

  renderer.render(scene,camera);
 }

 animate();

 if(introEnabled){
  setTimeout(()=>intro.classList.add("hide"),1200);
  setTimeout(()=>intro.remove(),3000);
 }else{
  intro.remove();
 }

 window.addEventListener("resize",()=>{
  const w=box.clientWidth;
  const h=box.clientHeight;
  camera.aspect=w/h;
  camera.updateProjectionMatrix();
  renderer.setSize(w,h);
 });

})();
</script>
</body>
</html>
"""

    html = html.replace("__NAME__", name)
    html = html.replace("__STYLE__", style)
    html = html.replace("__ANIMATION__", animation)
    html = html.replace("__INTRO__", intro)

    components.html(html, height=height, scrolling=False)


# ============================================================
# 6. SIDEBAR NAVIGATION
# ============================================================

with st.sidebar:
    st.markdown("# 💎 JON AI")
    st.caption("NEXT GENERATION INTELLIGENCE")
    st.markdown("---")

    selected_page = st.radio(
        "Navigatsiya",
        PAGES,
        index=PAGES.index(st.session_state.get("page", PAGES[0]))
        if st.session_state.get("page") in PAGES else 0
    )

    st.session_state.page = selected_page

    st.markdown("---")
    st.markdown("### SYSTEM STATUS")

    if get_api_key():
        st.success("API KEY: CONFIGURED")
    else:
        st.warning("API KEY: NOT CONFIGURED")

    st.caption(f"JON AI VERSION {APP_VERSION}")


# ============================================================
# 7. HOME
# ============================================================

def render_home():
    st.markdown(
        '<div class="hero-title">JON AI LIVE</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="hero-subtitle">'
        'NEXT GENERATION VIRTUAL INTELLIGENCE'
        '</div>',
        unsafe_allow_html=True
    )

    quote, support = st.session_state.motivation_quote
    st.markdown(
        f"""
        <div class="motivation-frame">
            <div class="motivation-label">✦ BUGUNGI MOTIVATSIYA ✦</div>
            <div class="motivation-quote">“{quote}”</div>
            <div class="motivation-subtext">{support}</div>
            <div class="motivation-signature">— JON AI</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    show_character(570)

    st.markdown("## ✨ JON AI imkoniyatlari")

    cards = [
        ("🤖 AI Chat", "Aqlli suhbat va yordam"),
        ("🎓 Ta'lim", "IELTS, ingliz tili va bilim"),
        ("📊 PowerPoint", "Taqdimot yaratish"),
        ("🌐 Tarjimon", "Ko'p tilli tarjima"),
        ("🎨 Creative Studio", "Ijodiy promptlar"),
        ("📈 Excel", "Ma'lumotlarni tahlil qilish")
    ]

    cols = st.columns(3)

    for i, (title, description) in enumerate(cards):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div class="feature-card">
                    <h3>{title}</h3>
                    <p>{description}</p>
                </div>
                """,
                unsafe_allow_html=True
            )


# ============================================================
# 8. AI CHAT
# ============================================================

def render_chat():
    st.title("🤖 JON AI — Intelligent Chat")

    if not get_api_key():
        st.warning("Avval Sozlamalar bo'limiga Gemini API kalitini kiriting.")

    c1, c2 = st.columns([5, 1])

    with c2:
        if st.button("🗑️ Tozalash"):
            st.session_state.messages = []
            st.rerun()

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    prompt = st.chat_input("JON AI ga savolingizni yozing...")

    if prompt:
        history = st.session_state.messages.copy()

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI fikrlamoqda..."):
                answer = ask_ai(prompt, history=history)
                st.markdown(answer)
                
        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })


# ============================================================
# 9. EDUCATION CENTER
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
            "Speaking",
            "Writing",
            "Reading",
            "Forex boshlang'ich ta'lim",
            "Dasturlash"
        ]
    )

    level = st.selectbox(
        "Daraja",
        [
            "Beginner",
            "Elementary",
            "Intermediate",
            "Upper-Intermediate",
            "Advanced"
        ]
    )

    task = st.selectbox(
        "Vazifa",
        [
            "Mavzuni tushuntir",
            "Dars tayyorla",
            "Test tuz",
            "Mashqlar yarat",
            "Savol-javob o'tkaz",
            "Imtihonga tayyorla"
        ]
    )

    extra = st.text_area("Qo'shimcha talab")

    if st.button("🎓 Darsni boshlash", use_container_width=True):
        prompt = f"""
        Yo'nalish: {topic}
        Daraja: {level}
        Vazifa: {task}
        Talab: {extra}

        Menga tushunarli dars tayyorla.
        Misollar va mashqlar bilan tushuntir.
        """

        with st.spinner("Dars tayyorlanmoqda..."):
            answer = ask_ai(prompt)

        st.markdown(answer)


# ============================================================
# 10. POWERPOINT GENERATOR
# ============================================================

def render_powerpoint():
    st.title("📊 Professional PowerPoint Studio")

    title = st.text_input("Taqdimot mavzusi")
    count = st.slider("Slaydlar soni", 3, 15, 7)
    language = st.selectbox(
        "Til",
        ["O'zbek", "English", "Русский"]
    )

    style = st.selectbox(
        "Dizayn",
        ["Professional", "Minimal", "Technology", "Education"]
    )

    if st.button("✨ Taqdimot yaratish", use_container_width=True):
        if not title.strip():
            st.warning("Mavzuni kiriting.")
            return

        prompt = f"""
        {title} mavzusida {count} ta slayd uchun taqdimot tayyorla.
        Til: {language}
        Uslub: {style}

        Javobni quyidagi formatda yoz:
        SLIDE 1: Sarlavha
        CONTENT: Mazmun

        Har bir slaydni alohida ajrat.
        """

        with st.spinner("Slaydlar tayyorlanmoqda..."):
            answer = ask_ai(prompt)

        st.markdown(answer)

        presentation = Presentation()
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)

        sections = re.split(
            r"SLIDE\s*\d+\s*:",
            answer,
            flags=re.IGNORECASE
        )

        sections = [s.strip() for s in sections if s.strip()]

        if not sections:
            sections = [answer]

        for section in sections[:count]:
            lines = section.splitlines()
            slide_title = lines[0][:100] if lines else title
            body = "\n".join(lines[1:]) if len(lines) > 1 else section

            slide = presentation.slides.add_slide(
                presentation.slide_layouts[6]
)
