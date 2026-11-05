
import streamlit as st
import pandas as pd
import json
import os
import base64
import re

from io import BytesIO
from google import genai
from pptx import Presentation

# ==================================
# 1. APP SOZLAMALARI
# ==================================

st.set_page_config(
    page_title="JON AI 3.5",
    page_icon="💎",
    layout="wide"
)

APP_VERSION = "3.5"
MODEL = "gemini-3.5-flash"

JON_CORE = """
Sen JON AI 3.5 nomli aqlli virtual yordamchisan.
O'zbek tilida ravon gapirasan.
Savollarga aniq va tushunarli javob berasan.
Ta'lim, tarjima, dasturlash va ijodiy ishlarda yordam berasan.
Doimo hurmat bilan muloqot qilasan.
"""

# ==================================
# 2. SESSION STATE
# ==================================

defaults = {
    "messages": [],
    "tasks": [],
    "character_name": "JON",
    "character_style": "Hologramma",
    "avatar": None,
    "api_key": ""
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ==================================
# 3. GEMINI AI
# ==================================

def ask_ai(prompt, extra_context=""):

    api_key = st.session_state.get(
        "api_key", ""
    ).strip()

    if not api_key:
        return "API kalitini Sozlamalardan kiriting."

    try:
        client = genai.Client(api_key=api_key)

        response = client.models.generate_content(
            model=MODEL,
            contents=(
                JON_CORE
                + "\n"
                + extra_context
                + "\n"
                + prompt
            )
        )

        if response.text:
            return response.text

        return "AI javob qaytarmadi."

    except Exception as e:
        return f"Gemini xatoligi: {str(e)}"


# ==================================
# 4. PERSONAJ
# ==================================


def show_character(answer=""):

    import streamlit as st
    import streamlit.components.v1 as components
    import json

    styles = {
        "Hologramma": "#00dfff",
        "Realistik robot": "#aabbd0",
        "Multfilm uslubi": "#ffb6d9",
        "Fantastik": "#a66cff"
    }

    if "character_style" not in st.session_state:
        st.session_state.character_style = "Hologramma"

    if "character_name" not in st.session_state:
        st.session_state.character_name = "JON"

    style = st.selectbox(
        "🎭 JON AI ko‘rinishi",
        list(styles.keys()),
        key="character_style"
    )

    color = styles[style]
    name = st.session_state.character_name

    html = """
    <div id="jon-panel">
      <div id="jon-status">● ONLINE</div>
      <canvas id="jon3d"></canvas>
      <h2>JON AI</h2>
      <p>Virtual yordamchingiz</p>
      <div id="jon-message">Salom! Men JON AI man.</div>
      <button id="speak">🔊 Gapirish</button>
      <button id="listen">🎙 Mikrofon</button>
      <div id="recognized">Ovozli buyruq shu yerda ko‘rinadi</div>
    </div>

    <style>
    #jon-panel {
      background:radial-gradient(circle,#12376c,#020817);
      color:white;text-align:center;
      border:2px solid COLOR;
      border-radius:25px;padding:20px;
    }
    #jon3d {
      width:100%;height:350px;
      display:block;
    }
    #jon-panel h2 {color:COLOR;font-size:30px}
    #jon-status {color:#00ff9d;font-weight:bold}
    #jon-message {
      background:#10264a;padding:15px;
      border-radius:15px;margin:15px;
    }
    #jon-panel button {
      background:#087edb;color:white;
      border:0;border-radius:12px;
      padding:12px;margin:5px;
    }
    #recognized {color:#9edfff;margin:12px}
    </style>

    <script src="https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.min.js"></script>

    <script>
    const color = "COLOR";
    const canvas = document.getElementById("jon3d");

    if (typeof THREE !== "undefined") {

      const scene = new THREE.Scene();

      const camera = new THREE.PerspectiveCamera(
        45, canvas.clientWidth / 350, 0.1, 100
      );
      camera.position.z = 5;

      const renderer = new THREE.WebGLRenderer({
        canvas:canvas, alpha:true, antialias:true
      });

      renderer.setPixelRatio(Math.min(window.devicePixelRatio,2));
      renderer.setSize(canvas.clientWidth,350);

      scene.add(new THREE.AmbientLight(0xffffff,2));

      const light = new THREE.PointLight(color,3);
      light.position.set(2,3,4);
      scene.add(light);

      const jon = new THREE.Group();
      scene.add(jon);

      const material = new THREE.MeshStandardMaterial({
        color:color,
        metalness:0.5,
        roughness:0.25
      });

      const head = new THREE.Mesh(
        new THREE.SphereGeometry(0.8,32,32),
        material
      );
      jon.add(head);

      const eyeMaterial = new THREE.MeshStandardMaterial({
        color:0xffffff,
        emissive:color,
        emissiveIntensity:2
      });

      const eyes = [];

      [-0.3,0.3].forEach(x => {
        const eye = new THREE.Mesh(
          new THREE.SphereGeometry(0.09,16,16),
          eyeMaterial
        );
        eye.position.set(x,0.12,0.7);
        jon.add(eye);
        eyes.push(eye);
      });

      const mouth = new THREE.Mesh(
        new THREE.BoxGeometry(0.3,0.04,0.04),
        new THREE.MeshBasicMaterial({color:color})
      );
      mouth.position.set(0,-0.35,0.72);
      jon.add(mouth);

      const body = new THREE.Mesh(
        new THREE.CylinderGeometry(0.45,0.7,1,32),
        material
      );
      body.position.y=-1.15;
      jon.add(body);

      const ring = new THREE.Mesh(
        new THREE.TorusGeometry(1.25,0.025,12,100),
        new THREE.MeshBasicMaterial({color:color})
      );
      ring.rotation.x=1.2;
      jon.add(ring);

      let talking=false;
      let blink=0;

      function animate() {
        requestAnimationFrame(animate);

        jon.rotation.y = Math.sin(Date.now()*0.0005)*0.15;
        jon.position.y = Math.sin(Date.now()*0.002)*0.08;

        ring.rotation.z += 0.008;

        if(talking) {
          mouth.scale.y = 1 + Math.abs(Math.sin(Date.now()*0.02))*7;
        } else {
          mouth.scale.y = 1;
        }

        blink++;

        if(blink % 300 < 5) {
          eyes.forEach(e => e.scale.y=0.1);
        } else {
          eyes.forEach(e => e.scale.y=1);
        }

        renderer.render(scene,camera);
      }

      animate();

      window.jonTalk = function(text) {
        if(!("speechSynthesis" in window)) {
          document.getElementById("jon-message").innerText =
            "Ovozli xizmat brauzerda qo‘llab-quvvatlanmaydi.";
          return;
        }

        window.speechSynthesis.cancel();

        const speech = new SpeechSynthesisUtterance(text);
        speech.lang="uz-UZ";
        speech.rate=1;
        speech.pitch=1;

        talking=true;

        speech.onend=()=>talking=false;
        speech.onerror=()=>talking=false;

        window.speechSynthesis.speak(speech);
      };

      document.getElementById("speak").onclick=()=>{
        window.jonTalk("Salom! Men JON AI. Sizga qanday yordam beray?");
      };

      document.getElementById("listen").onclick=()=>{
        const Recognition =
          window.SpeechRecognition ||
          window.webkitSpeechRecognition;

        if(!Recognition) {
          document.getElementById("recognized").innerText =
            "Mikrofon orqali tanish brauzeringizda mavjud emas.";
          return;
        }

        const recognition = new Recognition();
        recognition.lang="uz-UZ";
        recognition.interimResults=false;

        recognition.onresult=(event)=>{
          document.getElementById("recognized").innerText =
            "Siz: " + event.results[0][0].transcript;
        };

        recognition.onerror=()=>{
          document.getElementById("recognized").innerText =
            "Mikrofon ishlamadi. Ruxsatni tekshiring.";
        };

        recognition.start();
      };

    } else {
      document.getElementById("jon-message").innerText =
        "3D kutubxonasi yuklanmadi.";
    }
    </script>
    """

    html = html.replace("COLOR", color)

    components.html(html, height=570, scrolling=False)

    if answer:
        st.info(answer)

        if st.button("🔊 AI javobini eshitish", key="jon_ai_voice"):
            components.html(
                """
                <script>
                const text = """ + json.dumps(
                    str(answer), ensure_ascii=False
                ) + """;
                const speech = new SpeechSynthesisUtterance(text);
                speech.lang = "uz-UZ";
                speechSynthesis.cancel();
                speechSynthesis.speak(speech);
                </script>
                """,
                height=50
            )
            

def parse_slides(result):

    try:
        clean = result.strip()

        clean = re.sub(
            r"```(?:json)?",
            "",
            clean
        )

        clean = clean.replace("```", "").strip()

        start = clean.find("[")
        end = clean.rfind("]") + 1

        if start == -1 or end == 0:
            raise ValueError("JSON massiv topilmadi")

        data = json.loads(clean[start:end])

        if not isinstance(data, list):
            raise ValueError("Natija ro'yxat emas")

        return data

    except Exception as e:
        st.error(f"JSON xatosi: {e}")
        return []


# ==================================
# 6. POWERPOINT YARATISH
# ==================================

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

        points = item.get("points", [])

        if isinstance(points, str):
            points = [points]

        body.text = ""

        for point in points:

            paragraph = body.add_paragraph()
            paragraph.text = str(point)
            paragraph.level = 0

    output = BytesIO()

    prs.save(output)
    output.seek(0)

    return output


# ==================================
# 7. SIDEBAR
# ==================================

pages = [
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

st.sidebar.title("💎 JON AI 3.5")
st.sidebar.caption("Virtual Intelligence System")

page = st.sidebar.radio(
    "Bo'limni tanlang",
    pages
)

st.sidebar.divider()

st.sidebar.caption(
    "JON AI 3.5 | 2026"
)


# ==================================
# 8. BOSH SAHIFA
# ==================================

if page == "🏠 Bosh sahifa":

    st.title("💎 JON AI 3.5")

    st.caption(
        "Sizning shaxsiy virtual yordamchingiz"
    )

    show_character()

    st.markdown("### Xush kelibsiz!")

    st.write(
        "JON AI bilan suhbatlashing, bilim oling, "
        "tarjima qiling va yangi loyihalar yarating."
    )

    st.success("JON AI tizimi tayyor!")


# ==================================
# 9. AI CHAT
# ==================================

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

            with st.spinner("JON AI javob tayyorlamoqda..."):

                answer = ask_ai(prompt)

                st.markdown(answer)

        st.session_state.messages.append({
            "role": "assistant",
            "content": answer
        })

        st.rerun()


# ==================================
# 10. TA'LIM MARKAZI
# ==================================

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
            "Psixologiya",
            "Boshqa"
        ]
    )

    question = st.text_area(
        "Savolingiz yoki mavzuni yozing"
    )

    if st.button("📚 Tushuntirish"):

        if question:

            with st.spinner("Tayyorlanmoqda..."):

                answer = ask_ai(
                    question,
                    "Fan: " + subject +
                    ". Mavzuni o'quvchiga sodda tushuntir."
                )

                st.markdown(answer)

        else:
            st.warning("Avval savol yozing.")


# ==================================
# 11. POWERPOINT
# ==================================

elif page == "📊 PowerPoint":

    st.title("📊 JON AI — PowerPoint yaratuvchi")

    topic = st.text_input(
        "Taqdimot mavzusi"
    )

    slide_count = st.slider(
        "Slaydlar soni",
        3, 15, 6
    )

    if st.button("✨ Taqdimot yaratish"):

        if topic:

            prompt = (
                "Mavzu: " + topic +
                ". " + str(slide_count) +
                " ta slayd tayyorla. "
                "Faqat JSON massiv qaytar. "
                'Har bir elementda "title" va "points" bo\'lsin. '
                "points ro'yxat ko'rinishida bo'lsin."
            )

            with st.spinner("Taqdimot tayyorlanmoqda..."):

                result = ask_ai(prompt)

                slides_data = parse_slides(result)

                if slides_data:

                    pptx_file = create_presentation(
                        topic,
                        slides_data
                    )

                    st.success("Taqdimot tayyor!")

                    st.download_button(
                        "⬇️ PowerPoint yuklab olish",
                        data=pptx_file,
                        file_name="JON_AI_Taqdimot.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

        else:
            st.warning("Mavzuni kiriting.")


# ==================================
# 12. TARJIMON
# ==================================

elif page == "🌐 Tarjimon":

    st.title("🌐 JON AI — Universal tarjimon")

    source_lang = st.selectbox(
        "Matn qaysi tilda?",
        [
            "Avtomatik aniqlash",
            "O'zbek",
            "English",
            "Русский",
            "Turk",
            "Arab",
            "Nemis",
            "Fransuz"
        ]
    )

    target_lang = st.selectbox(
        "Qaysi tilga tarjima qilinsin?",
        [
            "O'zbek",
            "English",
            "Русский",
            "Turk",
            "Arab",
            "Nemis",
            "Fransuz"
        ]
    )

    text_to_translate = st.text_area(
        "Tarjima qilinadigan matn"
    )

    if st.button("🌐 Tarjima qilish"):

        if text_to_translate:

            prompt = (
                "Quyidagi matnni "
                + source_lang
                + " tilidan "
                + target_lang
                + " tiliga tarjima qil. "
                "Faqat tarjimani chiqar:\n"
                + text_to_translate
            )

            with st.spinner("Tarjima qilinmoqda..."):

                translated = ask_ai(prompt)

                st.markdown("### Natija")
                st.write(translated)

        else:
            st.warning("Matn kiriting.")


# ==================================
# 13. RASM PROMPTI
# ==================================

elif page == "🎨 Rasm prompti":

    st.title("🎨 JON AI — Rasm prompti")

    idea = st.text_area(
        "Qanday rasm yaratmoqchisiz?"
    )

    style = st.selectbox(
        "Rasm uslubi",
        [
            "Realistik",
            "3D",
            "Multfilm",
            "Anime",
            "Kiberpank",
            "Fantastik"
        ]
    )

    if st.button("✨ Prompt yaratish"):

        if idea:

            prompt = (
                "Quyidagi g'oya uchun batafsil "
                "rasm prompti yoz. Uslub: "
                + style
                + ". G'oya: "
                + idea
            )

            with st.spinner("Prompt yaratilmoqda..."):

                result = ask_ai(prompt)

                st.markdown("### Tayyor prompt")
                st.code(result)

                st.button(
                    "📋 Promptni nusxalash",
                    on_click=st.toast,
                    args=("Prompt tayyor!",)
                )

        else:
            st.warning("Rasm g'oyasini yozing.")


# ==================================
# 14. EXCEL TAHLIL
# ==================================

elif page == "📈 Excel tahlil":

    st.title("📈 JON AI — Excel tahlil")

    uploaded_file = st.file_uploader(
        "Excel yoki CSV fayl yuklang",
        type=["xlsx", "csv"]
    )

    if uploaded_file:

        try:

            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success("Fayl muvaffaqiyatli ochildi!")

            st.dataframe(
                df,
                use_container_width=True
            )

            st.markdown("### Asosiy ma'lumotlar")

            st.write(
                df.describe(include="all")
            )

            question = st.text_input(
                "Jadval haqida savolingiz"
            )

            if st.button("📊 Tahlil qilish"):

                if question:

                    context = df.head(30).to_string()

                    answer = ask_ai(
                        question,
                        "Jadval ma'lumotlari:\n" + context
                    )

                    st.markdown(answer)

                else:
                    st.warning("Savol kiriting.")

        except Exception as e:
            st.error(f"Fayl xatosi: {e}")


# ==================================
# 15. VAZIFALAR
# ==================================

elif page == "📝 Vazifalar":

    st.title("📝 JON AI — Kunlik vazifalar")

    new_task = st.text_input(
        "Yangi vazifa"
    )

    if st.button("➕ Vazifa qo'shish"):

        if new_task:

            st.session_state.tasks.append(new_task)
            st.rerun()

    st.markdown("### Vazifalar ro'yxati")

    for i, task in enumerate(
        st.session_state.tasks
    ):

        st.checkbox(
            task,
            key=f"task_{i}"
        )

    if st.button("🗑️ Barcha vazifalarni tozalash"):

        st.session_state.tasks = []
        st.rerun()


# ==================================
# 16. KAMERA
# ==================================

elif page == "📷 Kamera":

    st.title("📷 JON AI — Kamera")

    photo = st.camera_input(
        "Suratga olish"
    )

    if photo:

        st.image(
            photo,
            caption="Olingan surat",
            use_container_width=True
        )


# ==================================
# 17. PERSONAJ
# ==================================

elif page == "👤 Personaj":

    st.title("👤 JON AI — Personaj sozlamalari")

    st.text_input(
        "Personaj nomi",
        key="character_name"
    )

    st.selectbox(
        "Tashqi ko'rinish",
        [
            "Hologramma",
            "Robot",
            "Multfilm uslubi",
            "Fantastik"
        ],
        key="character_style"
    )

    avatar = st.file_uploader(
        "Personaj rasmi yuklash",
        type=["png", "jpg", "jpeg"]
    )

    if avatar:
        st.session_state.avatar = avatar.getvalue()

    if st.session_state.avatar:

        st.image(
            st.session_state.avatar,
            caption="Siz tanlagan personaj",
            use_container_width=True
        )

    else:
        show_character()

    st.info(
        "Personajning nomi, uslubi va rasmi "
        "shu bo'limda sozlanadi."
    )


# ==================================
# 18. SOZLAMALAR
# ==================================

elif page == "⚙️ Sozlamalar":

    st.title("⚙️ JON AI — Sozlamalar")

    st.markdown("### Gemini API")

    api_input = st.text_input(
        "Gemini API kaliti",
        type="password",
        value=st.session_state.api_key
    )

    if st.button("💾 API kalitini saqlash"):

        st.session_state.api_key = api_input.strip()

        st.success("API kaliti saqlandi!")

    if st.button("🔄 API ulanishini tekshirish"):

        if st.session_state.api_key:

            result = ask_ai(
                "Salom. Faqat OK deb javob ber."
            )

            if result.startswith("Gemini xatoligi:"):
                st.error(result)
            elif result == "API kalitini Sozlamalardan kiriting.":
                st.warning(result)
            else:
                st.success("API javob berdi!")
                st.write(result)

        else:
            st.warning("API kalitini kiriting!")

    st.divider()

    st.write("Ilova versiyasi:", APP_VERSION)
    st.write("Model:", MODEL)


# ==================================
# 19. PREMIUM
# ==================================

elif page == "💎 Premium":

    st.title("💎 JON AI Premium")

    st.subheader("Kelajakdagi imkoniyatlar")

    st.write("🚀 Kengaytirilgan AI imkoniyatlari")
    st.write("🎭 Maxsus personaj sozlamalari")
    st.write("🎙️ Ovozli yordamchi")
    st.write("✨ Maxsus dizayn va effektlar")
    st.write("📊 Kengaytirilgan tahlil")

    st.info(
        "Premium bo'limi hozircha namoyish uchun. "
        "To'lov tizimi hali ulanmagan."
    )


# ==================================
# 20. FOOTER
# ==================================

st.sidebar.divider()

st.sidebar.caption(
    "💎 JON AI — Your Virtual Assistant"
)

st.sidebar.caption(
    "Powered by Google Gemini"
)
