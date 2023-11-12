@echo off
start cmd /k "python app.py & streamlit run app.py streamlit run seu_script.py --server.port 8501 & start http://localhost:8501"