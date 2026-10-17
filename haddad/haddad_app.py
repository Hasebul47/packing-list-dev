"""Run alone:  streamlit run haddad_app.py"""
import streamlit as st
from haddad_ui import render

def run(file=None):
    render(file)

if __name__ == "__main__":
    st.set_page_config(page_title="Haddad packing list", layout="wide")
    run()

