import os
from dotenv import load_dotenv
import google.generativeai as genai

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")

if api_key:
    api_key = api_key.strip()
    print("Key loaded. Starts with:", api_key[:6])
else:
    print("No API key found.")
    exit()

genai.configure(api_key=api_key)

print("\nAvailable Gemini models for this key:\n")

for model in genai.list_models():
    if "generateContent" in model.supported_generation_methods:
        print(model.name)