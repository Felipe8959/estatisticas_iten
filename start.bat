@echo off
start cmd /k "python Z:\estatisticas_iten\app.py & streamlit run app.py --server.port 8501 & start http://localhost:8501"

