"""Run alone:  streamlit run tw_app.py"""
import streamlit as st
from tw_ui import render

def run(file=None):
    render(file)

if __name__ == "__main__":
    st.set_page_config(page_title="Tally Weijl packing list", layout="wide", page_icon="📦")
    run()

