"""
WSGI entry point для PythonAnywhere.
"""
import sys
import os

# Добавляем папку проекта в путь
path = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, path)

# Импортируем Flask-приложение
from web_app import app as application
