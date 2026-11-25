import ollama

def config_details(user_input):
        response = ollama.chat(
        model="qwen2.5",
        messages=[
            {"role": "user", "content": user_input}
        ])
        
        return response["message"]["content"]

def ai_response():
    
    while True:
        user_input =input("What do you wanna ask ? If you don't want to continue the conversation, Type end to exit the conversation: ")
        if user_input.lower() == "end":
            print("Exiting the conversation.")
            break
        
        else:
            response = config_details(user_input)
            print(response)


print(ai_response())