
import json
import re
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from pptx import Presentation


# ============================================================
# JON AI LIVE 7.0
# ============================================================

APP_VERSION = "7.0"
MODEL = "gemini-2.5-flash"

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
Asosiy til o'zbek tili.
Aniq, ravon va foydali javob ber.
Ma'lumotlarni to'qib chiqarma.
Ishonching komil bo'lmasa, buni ayt.
Murakkab vazifalarni bosqichma-bosqich tushuntir.
Maxfiy kalitlarni oshkor qilma.
"""


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "api_key": "",
    "page": PAGES[0]
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# DESIGN
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

.stApp {
    background:
    radial-gradient(ellipse at top left,
    rgba(20,100,170,.22),transparent 45%),
    linear-gradient(145deg,#07111f,#10152b);
    color:#edf5ff;
}
html, body, [class*="css"] {
    font-family:Inter,sans-serif;
}
[data-testid="stSidebar"] {
    background:#0b1729;
}
.block-container {
    max-width:1400px;
    padding-top:1.5rem;
}
.jon-hero {
    padding:28px;
    border:1px solid #28618a;
    border-radius:22px;
    background:linear-gradient(120deg,#123b61,#211d50);
}
.jon-hero h1 {
    color:#e7faff;
    font-size:42px;
}
.jon-card {
    background:rgba(20,48,75,.7);
    border:1px solid rgba(100,190,255,.2);
    border-radius:16px;
    padding:18px;
    min-height:125px;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# GEMINI
# ============================================================

def get_api_key():
    key = st.session_state.api_key.strip()

    if key:
        return key

    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return ""


def ask_ai(prompt, history=None):
    key = get_api_key()

    if not key:
        return "Gemini API kaliti topilmadi. Sozlamalar bo'limiga kiriting."

    try:
        client = genai.Client(api_key=key)

        conversation = ""

        for item in (history or [])[-12:]:
            conversation += (
                f"\n{item['role']}: {item['content']}"
            )

        response = client.models.generate_content(
            model=MODEL,
            contents=f"""
{JON_CORE}

Oldingi suhbat:
{conversation}

Yangi topshiriq:
{prompt}
"""
        )

        return response.text or "JON AI javob qaytarmadi."

    except Exception as e:
        return f"Gemini xatosi: {e}"


# ============================================================
# 3D CHARACTER
# ============================================================

def show_character(height=460):

    name = json.dumps(st.session_state.character_name)

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

    html = r"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <style>
    html,body {
        margin:0;
        width:100%;
        height:100%;
        overflow:hidden;
        background:#07101d;
    }
    #stage {
        width:100%;
        height:100%;
        position:relative;
        background:radial-gradient(ellipse,#19395d,#07101d 75%);
        border-radius:20px;
    }
    canvas {display:block;width:100%;height:100%}
    #label {
        position:absolute;
        bottom:18px;
        width:100%;
        text-align:center;
        color:#a9f5ff;
        font: bold 22px Arial;
        letter-spacing:6px;
        text-shadow:0 0 15px #00dfff;
    }
    </style>
    </head>
    <body>
    <div id="stage">
    <canvas id="scene"></canvas>
    <div id="label"></div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
    <script>
    (() => {
        const stage=document.getElementById('stage');
        const canvas=document.getElementById('scene');
        document.getElementById('label').textContent=__NAME__;

        if(!window.THREE) return;

        const T=window.THREE;
        const scene=new T.Scene();
        const camera=new T.PerspectiveCamera(
            40,stage.clientWidth/stage.clientHeight,.1,100
        );
        camera.position.z=7;

        const renderer=new T.WebGLRenderer({
            canvas,antialias:true,alpha:true
        });
        renderer.setPixelRatio(Math.min(devicePixelRatio,1.7));

        const color=__COLOR__;

        scene.add(new T.AmbientLight(0xffffff,1.4));

        const light=new T.PointLight(color,3,20);
        light.position.set(3,3,5);
        scene.add(light);

        const group=new T.Group();
        scene.add(group);

        const material=new T.MeshStandardMaterial({
            color:color,
            metalness:.45,
            roughness:.25,
            emissive:color,
            emissiveIntensity:.18
        });

        const dark=new T.MeshStandardMaterial({
            color:0x14243d,
            metalness:.6,
            roughness:.3
        });

        function sphere(parent,mat,pos,scale) {
            const obj=new T.Mesh(
                new T.SphereGeometry(1,32,24),mat
            );
            obj.position.set(...pos);
            obj.scale.set(...scale);
            parent.add(obj);
            return obj;
        }

        sphere(group,dark,[0,-1,0],[.85,1,.5]);

        const head=new T.Group();
        head.position.y=.65;
        group.add(head);

        sphere(head,material,[0,0,0],[.72,.85,.58]);

        const white=new T.MeshBasicMaterial({color:0xffffff});
        const black=new T.MeshBasicMaterial({color:0x06101c});

        sphere(head,white,[-.26,.08,.49],[.14,.17,.07]);
        sphere(head,white,[.26,.08,.49],[.14,.17,.07]);

        sphere(head,black,[-.25,.08,.55],[.07,.09,.035]);
        sphere(head,black,[.27,.08,.55],[.07,.09,.035]);

        const mouth=sphere(
            head,black,[0,-.32,.53],[.18,.035,.035]
        );

        const ring=new T.Mesh(
            new T.TorusGeometry(1.35,.015,8,100),
            new T.MeshBasicMaterial({
                color:color,transparent:true,opacity:.7
            })
        );
        ring.rotation.x=1.2;
        ring.position.y=-.5;
        group.add(ring);

        const clock=new T.Clock();

        function resize() {
            const w=stage.clientWidth;
            const h=stage.clientHeight;
            camera.aspect=w/h;
            camera.updateProjectionMatrix();
            renderer.setSize(w,h);
        }

        new ResizeObserver(resize).observe(stage);
        resize();

        function animate() {
            requestAnimationFrame(animate);

            const t=clock.getElapsedTime();

            group.position.y=Math.sin(t*1.4)*.08;
            group.rotation.y=Math.sin(t*.5)*.15;
            head.rotation.x=Math.sin(t*.8)*.04;
            ring.rotation.z+=.004;

            mouth.scale.y=.035+Math.abs(Math.sin(t*3))*.035;

            renderer.render(scene,camera);
        }

        animate();
    })();
    </script>
    </body>
    </html>
    """

    html = (
        html.replace("__NAME__", name)
        .replace("__COLOR__", color)
    )

    components.html(html,height=height,scrolling=False)


# ============================================================
# HOME
# ============================================================

def render_home():

    st.markdown("""
    <div class="jon-hero">
        <h1>💎 JON AI LIVE 7.0</h1>
        <p>Professional Virtual Intelligence System</p>
        <p>Ta'lim · Ijod · Tahlil · AI suhbat</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("JON — sizning virtual yordamchingiz")

    show_character(430)

    a,b,c=st.columns(3)

    with a:
        st.markdown("""
        <div class="jon-card">
        <h3>🤖 AI Intelligence</h3>
        Gemini AI yordamida suhbat.
        </div>
        """,unsafe_allow_html=True)

    with b:
        st.markdown("""
        <div class="jon-card">
        <h3>👤 Live Character</h3>
        3D virtual personaj.
        </div>
        """,unsafe_allow_html=True)

    with c:
        st.markdown("""
        <div class="jon-card">
        <h3>🚀 Creative Studio</h3>
        Ijodiy va ta'lim vositalari.
        </div>
        """,unsafe_allow_html=True)


# ============================================================
# AI CHAT
# ============================================================

def render_chat():

    st.title("🤖 JON AI — Chat")

    for item in st.session_state.messages:
        with st.chat_message(item["role"]):
            st.markdown(item["content"])

    prompt=st.chat_input("JON AI ga yozing...")

    if prompt:
        st.session_state.messages.append({
            "role":"user","content":prompt
        })

        with st.chat_message("user"):
            st.write(prompt)

        with st.chat_message("assistant"):
            with st.spinner("JON AI fikrlamoqda..."):
                answer=ask_ai(
                    prompt,
                    st.session_state.messages[:-1]
                )
            st.markdown(answer)

        st.session_state.messages.append({
            "role":"assistant","content":answer
        })


# ============================================================
# EDUCATION
# ============================================================

def render_education():

    st.title("🎓 Ta'lim markazi")

    topic=st.selectbox(
        "Yo'nalish",
        ["Ingliz tili","IELTS","Multilevel","Grammatika","Lug'at","Boshqa fan"]
    )

    mode=st.selectbox(
        "Yordam turi",
        ["Mavzuni tushuntirish","Test tuzish",
         "Javoblarni tekshirish","Reja tuzish","Mashq yaratish"]
    )

    level=st.selectbox(
        "Daraja",["Boshlang'ich","O'rta","Yuqori"]
    )

    question=st.text_area("Mavzu yoki topshiriq")

    if st.button("JON AI dan yordam olish"):
        if question.strip():
            result=ask_ai(
                f"Yo'nalish: {topic}. Yordam turi: {mode}. "
                f"Daraja: {level}. Topshiriq: {question}"
            )
            st.markdown(result)
        else:
            st.warning("Avval mavzu kiriting.")


# ============================================================
# POWERPOINT
# ============================================================

def render_powerpoint():

    st.title("📊 PowerPoint yaratish")

    topic=st.text_input("Taqdimot mavzusi")
    count=st.slider("Slaydlar soni",3,15,7)

    if st.button("Taqdimot yaratish"):
        if not topic.strip():
            st.warning("Mavzu kiriting.")
            return

        with st.spinner("Slaydlar tayyorlanmoqda..."):
            raw=ask_ai(
                f"""
                {topic} mavzusida {count} ta slayd yarat.
                Faqat JSON qaytar:
                [{{"title":"Sarlavha","points":["1","2","3"]}}]
                """
            )

        try:
            cleaned=re.sub(r"```(?:json)?|```","",raw).strip()
            data=json.loads(cleaned)

            prs=Presentation()

            for item in data:
                slide=prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text=str(item["title"])
                frame=slide.placeholders[1].text_frame
                frame.clear()

                for i,point in enumerate(item["points"]):
                    p=frame.paragraphs[0] if i==0 else frame.add_paragraph()
                    p.text=str(point)

            output=BytesIO()
            prs.save(output)

            st.download_button(
                "📥 PowerPoint yuklab olish",
                data=output.getvalue(),
                file_name="JON_AI_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"Taqdimot yaratishda xatolik: {e}")
            st.write(raw)


# ============================================================
# TRANSLATOR
# ============================================================

def render_translator():

    st.title("🌐 JON AI Tarjimon")

    source=st.text_area("Tarjima qilinadigan matn")

    language=st.selectbox(
        "Qaysi tilga?",
        ["O'zbek","Ingliz","Rus","Turk","Nemis","Fransuz"]
    )

    if st.button("Tarjima qilish"):
        if source.strip():
            st.markdown(
                ask_ai(f"Quyidagi matnni {language} tiliga tarjima qil: {source}")
            )


# ============================================================
# IMAGE PROMPT
# ============================================================

def render_image_prompt():

    st.title("🎨 AI Rasm prompti")

    idea=st.text_area("Qanday rasm yaratmoqchisiz?")

    style=st.selectbox(
        "Uslub",
        ["Realistik","3D","Anime","Kinematik","Minimalistik","Fantastik"]
    )

    if st.button("Prompt yaratish"):
        if idea.strip():
            result=ask_ai(
                f"{idea}. {style} uslubida professional image generation prompt yoz. Ingliz tilida."
            )
            st.code(result)


# ============================================================
# EXCEL
# ============================================================

def render_excel():

    st.title("📈 Excel tahlil")

    file=st.file_uploader(
        "Excel fayl yuklang",
        type=["xlsx","xls","csv"]
    )

    if file:
        try:
            if file.name.endswith(".csv"):
                df=pd.read_csv(file)
            else:
                df=pd.read_excel(file)

            st.dataframe(df,use_container_width=True)

            st.subheader("Statistik ma'lumot")
            st.write(df.describe(include="all"))

            if st.button("AI orqali tahlil"):
                st.write(
                    ask_ai(
                        "Quyidagi jadvalni tahlil qil:\n"+
                        df.head(30).to_string()
                    )
                )

        except Exception as e:
            st.error(f"Faylni o'qishda xatolik: {e}")


# ============================================================
# TASKS
# ============================================================

def render_tasks():

    st.title("📝 Vazifalar boshqaruvi")

    task=st.text_input("Yangi vazifa")

    if st.button("Vazifa qo'shish"):
        if task.strip():
            st.session_state.tasks.append({
                "title":task,
                "done":False
            })
            st.rerun()

    for i,item in enumerate(st.session_state.tasks):
        c1,c2=st.columns([5,1])

        with c1:
            done=st.checkbox(
                item["title"],
                value=item["done"],
                key=f"task_{i}"
            )
            item["done"]=done

        with c2:
            if st.button("🗑️",key=f"delete_{i}"):
                st.session_state.tasks.pop(i)
                st.rerun()


# ============================================================
# CAMERA
# ============================================================

def render_camera():

    st.title("📷 Kamera")

    photo=st.camera_input("Kamera orqali suratga oling")

    if photo:
        st.image(photo,caption="Yuklangan surat",use_container_width=True)
        st.info("Surat olindi. AI tahlili uchun keyingi bosqichda vision integratsiyasi qo'shiladi.")


# ============================================================
# CHARACTER
# ============================================================

def render_character():

    st.title("👤 JON virtual personaji")

    st.session_state.character_name=st.text_input(
        "Personaj nomi",
        value=st.session_state.character_name
    )

    st.session_state.character_style=st.selectbox(
        "Personaj ko'rinishi",
        ["Hologramma","Robot","Virtual inson","Multfilm uslubi","Fantastik"],
        index=[
            "Hologramma","Robot","Virtual inson","Multfilm uslubi","Fantastik"
        ].index(st.session_state.character_style)
    )

    show_character(520)


# ============================================================
# SETTINGS
# ============================================================

def render_settings():

    st.title("⚙️ JON AI sozlamalari")

    st.text_input(
        "Gemini API kaliti",
        type="password",
        key="api_key"
    )

    if st.button("API ulanishini tekshirish"):
        with st.spinner("Tekshirilmoqda..."):
            answer=ask_ai("Faqat OK deb javob ber.")
        if answer.startswith("Gemini xatosi"):
            st.error(answer)
        elif "kaliti topilmadi" in answer:
            st.warning(answer)
        else:
            st.success("Gemini API muvaffaqiyatli ulandi.")
            st.caption(answer[:200])

    st.caption("API kalitini GitHub kodiga yozmang.")


# ============================================================
# PREMIUM
# ============================================================

def render_premium():

    st.title("💎 JON AI Premium")

    st.markdown("""
    ### Premium imkoniyatlar rejalari

    - Professional AI video yaratish
    - AI rasm yaratish
    - Video montaj va sahnalar
    - Logo va brending dizayni
    - AI musiqa vositalari
    - Kengaytirilgan 3D personaj

    Bu imkoniyatlar uchun alohida API va xizmatlar integratsiyasi talab qilinishi mumkin.
    """)

    st.info("Premium xizmatlar hozircha rejalashtirilgan. Hali faollashtirilmagan.")


# ============================================================
# MAIN NAVIGATION
# ============================================================

def main():

    with st.sidebar:
        st.title("💎 JON AI")
        st.caption(f"LIVE VERSION {APP_VERSION}")
        st.divider()

        page=st.radio(
            "Navigatsiya",
            PAGES,
            key="page"
        )

        st.divider()
        st.caption("Professional Virtual Intelligence")

    if page==PAGES[0]:
        render_home()

    elif page==PAGES[1]:
        render_chat()

    elif page==PAGES[2]:
        render_education()

    elif page==PAGES[3]:
        render_powerpoint()

    elif page==PAGES[4]:
        render_translator()

    elif page==PAGES[5]:
        render_image_prompt()

    elif page==PAGES[6]:
        render_excel()

    elif page==PAGES[7]:
        render_tasks()

    elif page==PAGES[8]:
        render_camera()

    elif page==PAGES[9]:
        render_character()

    elif page==PAGES[10]:
        render_settings()

    elif page==PAGES[11]:
        render_premium()


# ============================================================
# START APPLICATION
# ============================================================

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error("JON AI ishga tushishida xatolik yuz berdi.")
        st.exception(e)
