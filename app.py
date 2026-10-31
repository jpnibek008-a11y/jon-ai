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
MODEL = "gemini-2.5-flash".          
JON_CORE = """
Sen JON AI nomli aqlli virtual yordamchisan.
O'zbek tilida ravon gapirasan.
Savollarga aniq va tushunarli javob berasan.
Ta'lim, tarjima, dasturlash va ijodiy ishlarda yordam berasan.
Foydalanuvchi faqat personajning tashqi ko'rinishini o'zgartiradi.
Sening asosiy xaraktering va shaxsiyating o'zgarmaydi.
"""
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
        st.session_state[def ask_ai(prompt, extra_context=""):
    api_key = st.session_state.get("api_key", "")
    if not api_key:
        return "Sozlamalardan Gemini API kalitini kiriting."
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=MODEL,
            contents=JON_CORE + "\n" + extra_context + "\n" + prompt
        )
        return response.text or "Javob olinmadi."
    except Exception as e:
        return f"Xatolik: {e}"
        def show_character():
    st.markdown(
        "<div style='text-align:center;padding:35px;"
        "border:1px solid #00d9ff;border-radius:20px;"
        "background:linear-gradient(135deg,#07152d,#17245c);'>"
        "<div style='font-size:100px;'>🤖</div>"
        "<h1 style='color:#00eaff;'>JON</h1>"
        "<p style='color:white;'>VIRTUAL INTELLIGENCE ONLINE</p>"
        "</div>",
        unsafe_allow_html=True
    )
    def get_image(query):
    try:
        url = "https://commons.wikimedia.org/w/api.php"
        params = {
            "action": "query",
            "generator": "search",
            "gsrsearch": query,
            "gsrnamespace": 6,
            "gsrlimit": 1,
            "prop": "imageinfo",
            "iiprop": "url",
            "format": "json"
        }
        r = requests.get(url, params=params, timeout=15)
        pages = r.json().get("query", {}).get("pages", {})
        for page in pages.values():
            image_url = page["imageinfo"][0]["url"]
            result = requests.get(image_url, timeout=20)
            if result.status_code == 200:
                return result.content
    except Exception:
        return None
    return None
