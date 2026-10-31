
import streamlit as st
from google import genai

# SAYT SOZLAMALARI
st.set_page_config(
    page_title="JON AI",
    page_icon="🤖",
    layout="centered"
)

# DIZAYN
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #050816, #10285c, #1e1b4b);
    color: white;
}

h1 {
    text-align: center;
    color: #00d9ff;
    text-shadow: 0 0 15px #007bff;
}

.subtitle {
    text-align: center;
    color: #cbd5e1;
    font-size: 18px;
}

.stChatMessage {
    border: 1px solid #00d9ff;
    border-radius: 15px;
    background: rgba(0, 30, 70, 0.35);
}

.stButton > button {
    background: linear-gradient(90deg, #007bff, #7b2cff);
    color: white;
    border: 1px solid #00d9ff;
    border-radius: 12px;
    padding: 10px 18px;
}

</style>
""", unsafe_allow_html=True)

# ASOSIY SARLAVHA
st.title("🤖 JON AI")

st.markdown(
    '<p class="subtitle">Sizning shaxsiy sun’iy intellektingiz</p>',
    unsafe_allow_html=True
)

st.divider()
# JON AI HOLOGRAMMA
st.markdown("""
<style>
.holo-stage {
    height: 300px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    position: relative;
    overflow: hidden;
    background: radial-gradient(circle, #123b70, #071329 70%);
    border: 1px solid #00d9ff;
    border-radius: 24px;
    box-shadow: 0 0 25px #007bff55;
}

.holo-head {
    width: 75px;
    height: 85px;
    border: 3px solid #00eaff;
    border-radius: 45% 45% 40% 40%;
    background: linear-gradient(135deg, #1c83bd55, #00d9ff22);
    box-shadow: 0 0 25px #00d9ff;
    position: relative;
    animation: float 3s ease-in-out infinite;
}

.holo-head:before {
    content: "◉   ◉";
    position: absolute;
    top: 25px;
    left: 13px;
    color: #aaffff;
    font-size: 18px;
    text-shadow: 0 0 12px #00ffff;
}

.holo-body {
    width: 145px;
    height: 85px;
    margin-top: 8px;
    border: 3px solid #00d9ff;
    border-radius: 65px 65px 15px 15px;
    background: linear-gradient(90deg, #0875aa55, #00eaff33);
    box-shadow: 0 0 25px #00d9ff;
    animation: float 3s ease-in-out infinite;
}

.holo-ring {
    width: 190px;
    height: 35px;
    border: 2px solid #00eaff;
    border-radius: 50%;
    box-shadow: 0 0 18px #00d9ff;
    margin-top: -8px;
}

.holo-name {
    color: #8ffaff;
    font-size: 19px;
    font-weight: bold;
    margin-top: 12px;
    text-shadow: 0 0 12px #00d9ff;
}

@keyframes float {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-10px); }
}
</style>

<div class="holo-stage">
    <div class="holo-head"></div>
    <div class="holo-body"></div>
    <div class="holo-ring"></div>
    <div class="holo-name">JON AI</div>
</div>
""", unsafe_allow_html=True)


# API KALITI
api_key = st.text_input(
    "🔑 Gemini API kalitini kiriting",
    type="password",
    help="API kalitingizni hech kimga yubormang."
)

# SUHBAT TARIXI
if "messages" not in st.session_state:
    st.session_state.messages = []

# OLDINGI XABARLARNI KO‘RSATISH
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# FOYDALANUVCHI SAVOLI
savol = st.chat_input("JON AI ga savolingizni yozing...")

if savol:

    st.session_state.messages.append({
        "role": "user",
        "content": savol
    })

    with st.chat_message("user"):
        st.markdown(savol)

    if not api_key:
        st.warning("Avval Gemini API kalitini kiriting.")

    else:
        try:
            client = genai.Client(api_key=api_key)

            with st.chat_message("assistant"):
                with st.spinner("JON AI o‘ylamoqda..."):

                    javob = client.models.generate_content(
                        model="gemini-3.5-flash-lite",
                        contents=[
                            "Sen JON AI nomli aqlli yordamchisan. "
                            "O‘zbek tilida aniq, samimiy va foydali javob ber.",
                            savol
                        ]
                    )

                    matn = javob.text or "Hozir javob yaratilmadi."

                    st.markdown(matn)

            st.session_state.messages.append({
                "role": "assistant",
                "content": matn
            })

        except Exception as e:
            st.error("Xatolik yuz berdi: " + str(e))

# YON PANEL
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    st.write("🤖 JON AI")
    st.write("🌐 O‘zbek tilidagi yordamchi")
    st.divider()

    if st.button("🗑️ Suhbatni tozalash"):
        st.session_state.messages = []
        st.rerun()

    st.caption("JON AI — sizning aqlli yordamchingiz.")
    
