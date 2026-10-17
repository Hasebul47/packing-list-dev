"""Run alone:  streamlit run justice_app.py"""
import streamlit as st
from justice_ui import render


def run(file=None):
    render(file)


if __name__ == "__main__":
    st.set_page_config(page_title="Justice packing list", layout="wide", page_icon="📦")
    run()
