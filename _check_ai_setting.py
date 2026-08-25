import store

print("ai_model_preference =", store.get_setting("ai_model_preference", "(未设置)"))
print("ollama_model =", store.get_setting("ollama_model", "(未设置)"))
print("ollama_url =", store.get_setting("ollama_url", "(未设置)"))
import os
print("ZHIPU_API_KEY set:", bool(os.getenv("ZHIPU_API_KEY")))
print("DEEPSEEK_API_KEY set:", bool(os.getenv("DEEPSEEK_API_KEY")))
