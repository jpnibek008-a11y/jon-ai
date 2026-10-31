import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from google import genai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(
    page_title="JON AI",
    page_icon="💎",
    layout="wide"
)

MODEL = "gemini-2.5-flash"

if "messages" not in st.session_state:
    st.session_state.messages = []

if "tasks" not in st.session_state:
    st.session_state.tasks = []

if "character_name" not in st.session_state:
    st.session_state.character_name = "JON"

if "character_style" not in st.session_state:
    st.session_state.character_style = "Hologram"

if "avatar" not in st.session_state:
    st.session_state.avatar = None

if "api_key" not in st.session_state:
    st.session_state.api_key = ""

st.markdown("""
<style>
.stApp {
    background: radial-gradient(
        circle at 50% 0%,
        #173d79 0%,
        #08152e 45%,
        #030711 100%
    );
    color: white;
}

h1, h2, h3, p, label {
    color: white !important;
}

section[data-testid="stSidebar"] {
    background: #071329;
    border-right: 1px solid #00d9ff;
}

.stButton > button {
    background: linear-gradient(90deg,#006eff,#762cff);
    color: white;
    border: 1px solid #00d9ff;
    border-radius: 13px;
    padding: 10px 20px;
    font-weight: bold;
    width: 100%;
}

.stButton > button:hover {
    border: 1px solid white;
    box-shadow: 0 0 18px #00d9ff;
}

.holo {
    text-align: center;
    padding: 25px;
    border: 1px solid #00d9ff;
    border-radius: 25px;
    background: rgba(0, 100, 255, 0.09);
    box-shadow: 0 0 35px rgba(0,217,255,.25);
}

.avatar-head {
    width: 95px;
    height: 95px;
    margin: auto;
    border-radius: 50%;
    background: radial-gradient(circle,#b9ffff,#00aaff,#173dff);
    box-shadow: 0 0 35px #00d9ff;
    animation: pulse 2s infinite alternate;
}

.avatar-body {
    width: 145px;
    height: 120px;
    margin: 15px auto;
    border-radius: 60% 60% 20% 20%;
    background: linear-gradient(140deg,#00d9ff,#1646bb,#772cff);
    box-shadow: 0 0 30px #007bff;
}

.ring {
    width: 180px;
    height: 35px;
    border: 3px solid #00d9ff;
    border-radius: 50%;
    margin: 15px auto;
    box-shadow: 0 0 20px #00d9ff;
}

@keyframes pulse {
    from {transform: scale(.95); opacity:.7;}
    to {transform: scale(1.08); opacity:1;}
}

.card {
    padding: 18px;
    border: 1px solid #168dff;
    border-radius: 16px;
    background: rgba(0,80,180,.15);
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)


def ask_ai(prompt):
    key = st.session_state.api_key

    if not key:
        return "Sozlamalar bo‘limiga Gemini API kalitini kiriting."

    try:
        client = genai.Client(api_key=key)

        response = client.models.generate_content(
            model=MODEL,
            contents=prompt
        )

        return response.text or "Javob olinmadi."

    except Exception as e:
        return f"AI xatosi: {e}"


def show_hologram():
    st.markdown('<div class="holo">', unsafe_allow_html=True)

    if st.session_state.avatar:
        st.image(st.session_state.avatar, width=180)
    else:
        st.markdown("""
        <div class="avatar-head"></div>
        <div class="avatar-body"></div>
        <div class="ring"></div>
        """, unsafe_allow_html=True)

    st.markdown(
        f"<h2>{st.session_state.character_name}</h2>",
        unsafe_allow_html=True
    )

    st.markdown(
        f"<p>✦ {st.session_state.character_style} SYSTEM ACTIVE ✦</p>",
        unsafe_allow_html=True
    )

    st.markdown("</div>", unsafe_allow_html=True)
