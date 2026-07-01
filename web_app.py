#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-приложение: Генератор актов приема-передачи техники
Автор: Артём Тамарян
"""

import os
import sys
import json
import webbrowser
import threading
from datetime import date
from flask import Flask, render_template, request, jsonify, send_from_directory

# Подключаем логику из app.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import generate_act, load_models, save_models, TEMPLATE_PATH, MODELS_FILE, OUTPUT_DIR, BASE_DIR

app = Flask(__name__)

# Маршрут для генерации PDF - заглушка
GENERATED_DIR = os.path.join(BASE_DIR, "generated_acts")


@app.route("/")
def index():
    """Главная страница."""
    models = load_models()
    today = date.today()
    months_ru = [
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря"
    ]
    return render_template("index.html",
                           models=models,
                           today_date=f'{today.day} {months_ru[today.month - 1]} {today.year} г.')


@app.route("/api/models", methods=["GET"])
def api_get_models():
    """Получить список моделей."""
    return jsonify(load_models())


@app.route("/api/models", methods=["POST"])
def api_add_model():
    """Добавить новую модель."""
    data = request.get_json()
    model = data.get("model", "").strip()
    if not model:
        return jsonify({"error": "Пустая модель"}), 400
    models = load_models()
    if model not in models:
        models.append(model)
        models.sort()
        save_models(models)
    return jsonify(models)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Сгенерировать акт."""
    data = request.get_json()

    name_arm = data.get("name_arm", "").strip()
    name_rus = data.get("name_rus", "").strip()
    gender = data.get("gender", "Мужской")
    citizenship = data.get("citizenship", "РФ")
    model = data.get("model", "").strip()
    serial = data.get("serial", "").strip()

    # Валидация
    errors = []
    if not name_arm:
        errors.append("Введите ФИО на армянском")
    elif len(name_arm.split()) < 3:
        errors.append("Армянское ФИО должно содержать 3 части (Անուն Հայրանուն Ազգանուն)")

    if not name_rus:
        errors.append("Введите ФИО на русском")
    elif len(name_rus.split()) < 3:
        errors.append("Русское ФИО должно содержать 3 части (Фамилия Имя Отчество)")

    if not model:
        errors.append("Введите модель ноутбука")

    if not serial:
        errors.append("Введите серийный номер")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    # Создаём папку
    rus_parts = name_rus.split()
    folder_name = f"{rus_parts[0]}_{date.today().strftime('%Y-%m-%d')}"
    output_subdir = os.path.join(OUTPUT_DIR, folder_name)
    os.makedirs(output_subdir, exist_ok=True)

    filename = f"Акт_приема_передачи_{rus_parts[0]}_{date.today().strftime('%Y-%m-%d')}.docx"
    output_path = os.path.join(output_subdir, filename)

    try:
        generate_act(name_arm, name_rus, gender, citizenship, model, serial, output_path)
    except Exception as e:
        return jsonify({"success": False, "errors": [f"Ошибка генерации: {str(e)}"]}), 500

    # Сохраняем модель
    models = load_models()
    if model not in models:
        models.append(model)
        models.sort()
        save_models(models)

    return jsonify({
        "success": True,
        "file_path": output_path,
        "filename": filename,
        "folder": folder_name
    })


@app.route("/api/download/<folder>/<filename>")
def api_download(folder, filename):
    """Скачать сгенерированный файл."""
    directory = os.path.join(GENERATED_DIR, folder)
    return send_from_directory(directory, filename, as_attachment=True)


def open_browser():
    """Открыть браузер через 1.5 сек после запуска (только на Windows)."""
    if sys.platform == "win32":
        threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5678")).start()


if __name__ == "__main__":
    print("🌐 Запуск веб-приложения...")
    print(f"📁 Шаблон: {TEMPLATE_PATH}")
    print(f"📁 Модели: {MODELS_FILE}")
    print(f"📁 Сохранение: {OUTPUT_DIR}")


if __name__ == "__main__":
    print("🌐 Запуск веб-приложения...")
    print(f"📁 Шаблон: {TEMPLATE_PATH}")
    print(f"📁 Модели: {MODELS_FILE}")
    print(f"📁 Сохранение: {OUTPUT_DIR}")
    print(f"\n🔗 Открыть: http://127.0.0.1:5678")
    print("❌ Закрыть: Ctrl+C в терминале\n")

    # Создаём папку для генерации
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    open_browser()
    app.run(host="127.0.0.1", port=5678, debug=False)
