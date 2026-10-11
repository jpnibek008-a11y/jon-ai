import streamlit as st
from google import genai

st.set_page_config(
    page_title="JON AI",
    page_icon="🤖",
    layout="centered"
)

st.title("🤖 JON AI")
st.caption("Sizning shaxsiy sun'iy intellektingiz")

api_key = st.text_input(
    "Gemini API kalitingizni kiriting",
    type="password"
)

if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

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
                javob = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=savol
                )
                st.markdown(javob.text)

            st.session_state.messages.append({
                "role": "assistant",
                "content": javob.text
            })

        except Exception as e:
            st.error(f"Xatolik yuz berdi: {e}")
