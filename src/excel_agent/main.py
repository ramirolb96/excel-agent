import json
import os
import sys

# Ensure we can find the sibling script
sys.path.append(os.path.join(os.path.dirname(__file__), "..", ".."))

from dotenv import load_dotenv
from openai import OpenAI
from src.excel_agent.add_next_month import add_next_month

# 1. Load Environment Variables
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# 2. Define the "Tools"
tools = [
    {
        "type": "function",
        "function": {
            "name": "create_next_month_tab",
            "description": "Creates a new month tab in the Excel budget. It copies the previous month, clears variable expenses, updates dates, and handles installment logic.",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "The full path to the excel file. OPTIONAL: If the user does not specify a file, do NOT ask for it. Leave this blank.",
                    }
                },
                "required": [],
            },
        },
    }
]


# 3. The Agent Logic
def run_agent():
    print("🤖 Excel Agent Online. Type 'quit' to exit.")
    print("------------------------------------------")

    # Your Hardcoded Default Path
    DEFAULT_PATH = "/Users/ramirolb/Library/CloudStorage/OneDrive-Personal/Excel Documents/monthly-expenses.xlsx"

    # UPDATED SYSTEM PROMPT: We strictly tell the AI to use the default path
    messages = [
        {
            "role": "system",
            "content": (
                "You are an intelligent assistant that manages the user's Excel budget file. "
                f"You have a default file path configured at: {DEFAULT_PATH}. "
                "Unless the user explicitly asks to use a *different* file, NEVER ask for the file path. "
                "Just call the tool immediately."
            ),
        }
    ]

    while True:
        try:
            user_input = input("\nYou: ")
        except KeyboardInterrupt:
            break

        if user_input.lower() in ["quit", "exit"]:
            break

        messages.append({"role": "user", "content": user_input})

        response = client.chat.completions.create(
            model="gpt-4o", messages=messages, tools=tools, tool_choice="auto"
        )

        response_message = response.choices[0].message

        if response_message.tool_calls:
            print("🧠 AI Thinking: Using default path...")
            messages.append(response_message)

            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                function_args = json.loads(tool_call.function.arguments)

                if function_name == "create_next_month_tab":
                    # Use the arg if provided, otherwise use DEFAULT_PATH
                    path_to_use = function_args.get("file_path", DEFAULT_PATH)

                    print(f"🔧 Executing Python Script on: {path_to_use}...")

                    try:
                        add_next_month(path_to_use)
                        tool_output = "Success: Created the new month tab."
                    except Exception as e:
                        tool_output = f"Error: {str(e)}"

                    messages.append(
                        {
                            "role": "tool",
                            "tool_call_id": tool_call.id,
                            "name": function_name,
                            "content": tool_output,
                        }
                    )

            final_response = client.chat.completions.create(
                model="gpt-4o", messages=messages
            )
            ai_reply = final_response.choices[0].message.content
            print(f"🤖 Agent: {ai_reply}")
            messages.append({"role": "assistant", "content": ai_reply})

        else:
            print(f"🤖 Agent: {response_message.content}")
            messages.append(response_message)


if __name__ == "__main__":
    run_agent()
