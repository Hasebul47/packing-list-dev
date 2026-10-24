"""Run alone:  streamlit run anf_app.py"""
import streamlit as st
from anf_ui import render

def run(pdf_file=None):
    render(pdf_file)

if __name__ == "__main__":
    st.set_page_config(page_title="A&F packing list", layout="wide")
    run()

