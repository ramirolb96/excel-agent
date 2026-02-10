import json
import os
import sys
from datetime import datetime

# Ensure we can find sibling scripts
sys.path.append(os.path.join(os.path.dirname(__file__), "..", ".."))

from dotenv import load_dotenv
from openai import OpenAI

# IMPORT YOUR TOOLS
from src.excel_agent.add_next_month import add_next_month
from src.excel_agent.log_expense import log_expense

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# --- TOOL DEFINITIONS ---
tools = [
    {
        "type": "function",
        "function": {
            "name": "create_next_month_tab",
            "description": "Creates a new month tab in the Excel budget. Copies previous month, clears variables, updates dates.",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Optional file path.",
                    }
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "log_expense",
            "description": "Logs a new expense into the current month's Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "expense_name": {
                        "type": "string",
                        "description": "The name of the expense (e.g. 'Grocery Store', 'Uber').",
                    },
                    "amount": {
                        "type": "number",
                        "description": "The dollar amount spent.",
                    },
                    "date": {
                        "type": "string",
                        "description": "The date of the expense in YYYY-MM-DD format. Defaults to today if not specified.",
                    },
                },
                "required": ["expense_name", "amount"],
            },
        },
    },
]


def run_agent():
    print("🤖 Excel Agent Online. Type 'quit' to exit.")
    print("------------------------------------------")

    DEFAULT_PATH = "/Users/ramirolb/Library/CloudStorage/OneDrive-Personal/Excel Documents/monthly-expenses.xlsx"

    # System Prompt: Tell the AI what day it is so it can log dates correctly!
    today_str = datetime.now().strftime("%Y-%m-%d")
    messages = [
        {
            "role": "system",
            "content": (
                f"You are a helpful budget assistant. Today's date is {today_str}. "
                f"Your default Excel file is at: {DEFAULT_PATH}. "
                "Always use this default path unless the user provides a different one."
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

        # 1. Ask AI
        response = client.chat.completions.create(
            model="gpt-4o", messages=messages, tools=tools, tool_choice="auto"
        )

        response_message = response.choices[0].message

        # 2. Check for Tool Calls
        if response_message.tool_calls:
            print("🧠 AI is thinking...")
            messages.append(response_message)

            for tool_call in response_message.tool_calls:
                function_name = tool_call.function.name
                args = json.loads(tool_call.function.arguments)
                tool_output = ""

                # --- TOOL 1: Create Month ---
                if function_name == "create_next_month_tab":
                    path = args.get("file_path", DEFAULT_PATH)
                    print("🔧 Tool: Creating New Month...")
                    try:
                        add_next_month(path)
                        tool_output = "Success: Created new month tab."
                    except Exception as e:
                        tool_output = f"Error: {str(e)}"

                # --- TOOL 2: Log Expense ---
                elif function_name == "log_expense":
                    path = DEFAULT_PATH  # Always use default unless complex logic added
                    name = args.get("expense_name")
                    amt = args.get("amount")
                    date = args.get("date")  # Optional

                    print(f"🔧 Tool: Logging '{name}' (${amt})...")
                    try:
                        tool_output = log_expense(path, name, amt, date)
                    except Exception as e:
                        tool_output = f"Error: {str(e)}"

                # Send Result back to AI
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "name": function_name,
                        "content": tool_output,
                    }
                )

            # 3. Get Final Reply
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
