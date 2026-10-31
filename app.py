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
    background: linear-gradient(135deg, #0f172a, #1e1b4b);
    color: white;
}
h1 {
    text-align: center;
    color: #a78bfa;
}
.subtitle {
    text-align: center;
    color: #cbd5e1;
    font-size: 18px;
}
.stChatMessage {
    border-radius: 15px;
}
</style>
""", unsafe_allow_html=True)

# SARLAVHA
st.title("🤖 JON AI")
st.markdown(
    '<p class="subtitle">Sizning shaxsiy sun’iy intellektingiz</p>',
    unsafe_allow_html=True
)

st.divider()

# API KALITI
api_key = st.text_input(
    "🔑 Gemini API kalitini kiriting",
    type="password",
    help="API kalitingizni boshqalarga bermang."
)

# SUHBAT TARIXI
if "messages" not in st.session_state:
    st.session_state.messages = []

# OLDINGI XABARLARNI CHIQARISH
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# XABAR YOZISH
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
                        model="gemini-3.8-flash",
                    ,
                        contents=[
                            "Sen JON AI nomli yordamchisan. "
                            "O‘zbek tilida aniq va samimiy javob ber.",
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
            st.error(f"Xatolik yuz berdi: {e}")

# YON PANEL
with st.sidebar:
    st.header("⚙️ Sozlamalar")
    st.write("🤖 JON AI")
    st.write("🌐 O‘zbek tilidagi yordamchi")

    if st.button("🗑️ Suhbatni tozalash"):
        st.session_state.messages = []
        st.rerun()

    st.caption("JON AI — sizning aqlli yordamchingiz.")
