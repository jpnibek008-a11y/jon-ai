
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
from pptx.util import Inches, Pt


# ============================================================
# 1. CONFIGURATION
# ============================================================

APP_VERSION = "9.0"
DEFAULT_MODEL = "gemini-3.8-flash"

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
    "last_answer": ""
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


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
 const introEnabled=__INTRO__;

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

 for(let i=0;i<count;i++){
  positions[i*3]=(Math.random()-0.5)*15;
  positions[i*3+1]=(Math.random()-0.5)*9;
  positions[i*3+2]=(Math.random()-0.5)*9;
 }

 const geometry=new THREE.BufferGeometry();
 geometry.setAttribute(
  "position",new THREE.BufferAttribute(positions,3)
 );

 const particles=new THREE.Points(
  geometry,
  new THREE.PointsMaterial({
   color:0x72caff,size:0.035,
   transparent:true,opacity:0.8
  })
 );
 scene.add(particles);

 const clock=new THREE.Clock();

 function animate(){
  requestAnimationFrame(animate);
  const t=clock.getElapsedTime();

  [human,android].forEach((p,i)=>{
   const phase=i*Math.PI;

   p.root.position.y=Math.sin(t*1.4+phase)*0.12;
   p.body.rotation.z=Math.sin(t*1.1+phase)*0.035;
   p.head.rotation.y=Math.sin(t*0.8+phase)*0.12;
   p.head.rotation.x=Math.sin(t*1.5+phase)*0.035;

   p.armL.rotation.z=
    Math.sin(t*2.1+phase)*0.35+0.1;

   p.armR.rotation.z=
    -Math.sin(t*2.1+phase)*0.35-0.1;

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

            title_box = slide.shapes.add_textbox(
                Inches(0.7), Inches(0.5),
                Inches(12), Inches(1.0)
            )

            tf = title_box.text_frame
            tf.text = slide_title

            for paragraph in tf.paragraphs:
                paragraph.font.size = Pt(28)
                paragraph.font.bold = True

            body_box = slide.shapes.add_textbox(
                Inches(0.8), Inches(1.8),
                Inches(11.8), Inches(4.8)
            )

            body_box.text_frame.text = body[:3000]

            for paragraph in body_box.text_frame.paragraphs:
                paragraph.font.size = Pt(20)

        output = BytesIO()
        presentation.save(output)
        output.seek(0)

        st.download_button(
            "📥 PowerPoint yuklab olish",
            data=output,
            file_name="JON_AI_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )


# ============================================================
# 11. TRANSLATOR
# ============================================================

def render_translator():
    st.title("🌐 JON AI Translator")

    source = st.selectbox(
        "Asl til",
        [
            "Avtomatik",
            "O'zbek",
            "English",
            "Русский",
            "Turkish",
            "Arabic"
        ]
    )

    target = st.selectbox(
        "Tarjima tili",
        [
            "O'zbek",
            "English",
            "Русский",
            "Turkish",
            "Arabic",
            "German",
            "French"
        ]
    )

    text = st.text_area("Tarjima qilinadigan matn", height=180)

    if st.button("🌐 Tarjima qilish", use_container_width=True):
        if not text.strip():
            st.warning("Matn kiriting.")
            return

        prompt = f"""
        Quyidagi matnni {source} tilidan {target} tiliga tarjima qil.
        Ma'noni saqla, grammatik jihatdan to'g'ri tarjima qil.

        MATN:
        {text}
        """

        with st.spinner("Tarjima qilinmoqda..."):
            result = ask_ai(prompt)

        st.markdown("### Natija")
        st.text_area("Tarjima", result, height=200)


# ============================================================
# 12. IMAGE PROMPT STUDIO
# ============================================================

def render_image_prompt():
    st.title("🎨 AI Image Prompt Studio")

    subject = st.text_input("Rasm mavzusi")

    style = st.selectbox(
        "Tasvir uslubi",
        [
            "Photorealistic",
            "3D cinematic",
            "Anime",
            "Cyberpunk",
            "Fantasy",
            "Minimalist",
            "Luxury advertising"
        ]
    )

    lighting = st.selectbox(
        "Yoritish",
        [
            "Cinematic lighting",
            "Neon lighting",
            "Natural light",
            "Studio lighting"
        ]
    )

    ratio = st.selectbox(
        "Format",
        ["16:9", "1:1", "9:16", "4:5"]
    )

    details = st.text_area("Qo'shimcha detallar")

    if st.button("✨ Prompt yaratish", use_container_width=True):
        prompt = f"""
        Create a professional image generation prompt.
        Subject: {subject}
        Style: {style}
        Lighting: {lighting}
        Aspect ratio: {ratio}
        Details: {details}

        Write a detailed English prompt and a negative prompt.
        """

        with st.spinner("Prompt yaratilmoqda..."):
            result = ask_ai(prompt)

        st.markdown(result)
        st.session_state.copied_prompt = result
        st.code(result, language="text")


# ============================================================
# 13. EXCEL ANALYSIS
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

            st.success("Fayl muvaffaqiyatli yuklandi.")

            st.markdown("### Ma'lumotlar")
            st.dataframe(df, use_container_width=True)

            c1, c2, c3 = st.columns(3)

            c1.metric("Qatorlar", len(df))
            c2.metric("Ustunlar", len(df.columns))
            c3.metric(
                "Bo'sh qiymatlar",
                int(df.isna().sum().sum())
            )

            st.markdown("### Statistik ma'lumotlar")

            st.dataframe(
                df.describe(include="all").transpose(),
                use_container_width=True
            )

            question = st.text_area(
                "JON AI dan qanday tahlil kerak?"
            )

            if st.button("📊 AI tahlil"):
                context = df.head(30).to_string()
                context += "\n\nStatistika:\n"
                context += df.describe(include="all").to_string()

                prompt = question or (
                    "Ushbu jadvalni professional tahlil qil. "
                    "Muhim xulosalar va tavsiyalar ber."
                )

                with st.spinner("Tahlil qilinmoqda..."):
                    result = ask_ai(
                        prompt,
                        extra_context=context
                    )

                st.markdown(result)

        except Exception as error:
            st.error(f"Faylni o'qishda xato: {error}")


# ============================================================
# 14. TASK MANAGER
# ============================================================

def render_tasks():
    st.title("📝 JON AI Task Manager")

    with st.form("task_form", clear_on_submit=True):
        task = st.text_input("Yangi vazifa")

        priority = st.selectbox(
            "Muhimlik",
            ["Oddiy", "Muhim", "Juda muhim"]
        )

        submitted = st.form_submit_button("➕ Vazifa qo'shish")

        if submitted and task.strip():
            st.session_state.tasks.append({
                "task": task,
                "priority": priority,
                "done": False,
                "created": time.strftime("%Y-%m-%d %H:%M")
            })
            st.rerun()

    if not st.session_state.tasks:
        st.info("Hozircha vazifalar mavjud emas.")

    for i, item in enumerate(st.session_state.tasks):
        c1, c2, c3 = st.columns([5, 2, 1])

        with c1:
            st.write(
                ("✅ " if item["done"] else "⬜ ")
                + item["task"]
            )

        with c2:
            st.caption(item["priority"])

        with c3:
            if st.button("🗑️", key=f"delete_{i}"):
                st.session_state.tasks.pop(i)
                st.rerun()

        new_status = st.checkbox(
            "Bajarildi",
            value=item["done"],
            key=f"done_{i}"
        )

        if new_status != item["done"]:
            item["done"] = new_status


# ============================================================
# 15. CAMERA
# ============================================================

def render_camera():
    st.title("📷 JON AI Camera Studio")

    st.info(
        "Kameradan foydalanish uchun brauzer yoki "
        "qurilmangiz ruxsatini bering."
    )

    photo = st.camera_input("Suratga olish")

    if photo:
        st.image(
            photo,
            caption="Sizning suratingiz",
            use_container_width=True
        )

        st.success("Surat olindi.")

        if st.button("🤖 Rasmni AI bilan tahlil qilish"):
            st.info(
                "Kamera orqali surat olish mavjud. "
                "AI rasm tahlili uchun Gemini Vision "
                "integratsiyasi keyingi bosqichda qo'shiladi."
            )


# ============================================================
# 16. CHARACTER STUDIO
# ============================================================

def render_character():
    st.title("👤 JON AI — Character Studio")

    show_character(520, cinematic=False)

    st.markdown("### Personaj sozlamalari")

    name = st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    styles = [
        "Hologramma",
        "Realistik robot",
        "Multfilm uslubi",
        "Fantastik"
    ]

    style = st.selectbox(
        "Personaj uslubi",
        styles,
        index=styles.index(st.session_state.character_style)
        if st.session_state.character_style in styles else 0
    )

    animation = st.selectbox(
        "Animatsiya rejimi",
        ["Raqs", "Tabiiy harakat", "Sokin", "Energetik"]
    )

    intro = st.checkbox(
        "Kirish effektini yoqish",
        value=st.session_state.intro_enabled
    )

    if st.button("💾 Saqlash"):
        st.session_state.character_name = name
        st.session_state.character_style = style
        st.session_state.character_animation = animation
        st.session_state.intro_enabled = intro

        st.success("Personaj sozlamalari saqlandi.")
        st.rerun()

    st.caption(
        "Haqiqiy yuz ifodalari, ovoz bilan lab sinxronlash "
        "va kamera orqali kuzatish alohida integratsiyani talab qiladi."
    )


# ============================================================
# 17. SETTINGS
# ============================================================

def render_settings():
    st.title("⚙️ JON AI — System Settings")

    st.markdown("### Gemini API")

    key = st.text_input(
        "Gemini API Key",
        type="password",
        value=st.session_state.api_key
    )

    models = [
    "gemini-3.8-flash",
    "gemini-3.5-flash-lite"
    ]

    current_model = st.session_state.model

    model = st.selectbox(
        "Gemini modeli",
        models,
        index=models.index(current_model)
        if current_model in models else 0
    )

    if st.button("💾 API sozlamalarini saqlash"):
        st.session_state.api_key = key.strip()
        st.session_state.model = model
        st.success("Sozlamalar saqlandi.")

    st.markdown("---")
    st.markdown("### Ulanishni tekshirish")

    if st.button("🔍 Gemini API ni tekshirish"):
        if not get_api_key():
            st.error("Avval API kalitini kiriting.")
        else:
            with st.spinner("Ulanish tekshirilmoqda..."):
                if check_model():
                    st.success("Gemini API muvaffaqiyatli ishlayapti.")

    st.markdown("---")

    st.markdown("### System information")
    st.write("JON AI Version:", APP_VERSION)
    st.write("Selected model:", st.session_state.model)

    st.warning(
        "API kalitni ommaviy GitHub kodiga yozmang. "
        "Streamlit Secrets orqali saqlash xavfsizroq."
    )


# ============================================================
# 18. PREMIUM CREATIVE STUDIO
# ============================================================

def render_premium():
    st.title("💎 JON AI — Premium Creative Studio")

    st.markdown("""
    <div class="panel">
    <h2>🚀 NEXT GENERATION CREATIVE LAB</h2>
    Professional ijodiy loyihalar uchun yagona markaz.
    </div>
    """, unsafe_allow_html=True)

    features = [
        ("🎬 AI Video Studio", "Video ssenariy va sahnalar"),
        ("🎵 AI Music Studio", "Original musiqa g'oyalari"),
        ("🏢 Logo Studio", "Kompaniya logotipi konsepsiyasi"),
        ("🎭 Character Animation", "Personaj animatsiyasi"),
        ("📸 Photo to Movie", "Suratdan video ssenariysi"),
        ("🎨 Photo to Cartoon", "Multfilm uslubidagi promptlar"),
        ("🎞️ Movie Maker", "Film ssenariysi va storyboard"),
        ("✨ Visual Effects", "Maxsus effektlar g'oyalari")
    ]

    cols = st.columns(2)

    for i, (title, description) in enumerate(features):
        with cols[i % 2]:
            st.markdown(
                f"""
                <div class="feature-card">
                    <h3>{title}</h3>
                    <p>{description}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("---")
    st.markdown("### Creative AI Assistant")

    project = st.selectbox(
        "Loyiha turi",
        [
            "Kompaniya logotipi",
            "Reklama videosi",
            "Kino ssenariysi",
            "Musiqa konsepsiyasi",
            "3D personaj",
            "Animatsion video",
            "Professional prezentatsiya"
        ]
    )

    idea = st.text_area("Loyiha g'oyasi")

    if st.button("✨ Creative loyiha yaratish"):
        if not idea.strip():
            st.warning("Loyiha g'oyasini kiriting.")
            return

        prompt = f"""
        Professional creative director sifatida ishlang.

        Loyiha turi: {project}
        G'oya: {idea}

        Professional konsepsiya, vizual yo'nalish,
        ranglar, sahnalar va amalga oshirish rejasini yozing.
        """

        with st.spinner("Creative loyiha yaratilmoqda..."):
            result = ask_ai(prompt)

        st.markdown(result)


# ============================================================
# 19. MAIN ROUTER
# ============================================================

page = st.session_state.page

if page == "🏠 Bosh sahifa":
    render_home()

elif page == "🤖 AI Chat":
    render_chat()

elif page == "🎓 Ta'lim markazi":
    render_education()

elif page == "📊 PowerPoint":
    render_powerpoint()

elif page == "🌐 Tarjimon":
    render_translator()

elif page == "🎨 Rasm prompti":
    render_image_prompt()

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
# 20. FOOTER
# ============================================================

st.markdown("---")

st.markdown(
    """
    <div style="text-align:center;color:#7895bd;padding:12px">
        💎 JON AI LIVE 9.0<br>
        NEXT GENERATION VIRTUAL INTELLIGENCE<br>
        <small>Designed for the future.</small>
    </div>
    """,
    unsafe_allow_html=True
            )
    
