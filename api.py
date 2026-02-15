from flask import Flask, request, Response
import streamlit as st
import sys
import os

# यह Vercel के लिए adapter है
app = Flask(__name__)

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def catch_all(path):
    # Streamlit app को run करने के लिए
    os.system(f'streamlit run app.py --server.port={os.environ.get("PORT", 8501)} --server.enableCORS=false --server.enableXsrfProtection=false')
    return Response("Streamlit app is starting...")

# Vercel के लिए handler
def handler(request):
    return app(request)
