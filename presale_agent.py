import os
import shutil
import glob
import json
import chainlit as cl
from openai import OpenAI
from dotenv import load_dotenv
import pandas as pd
from pandasql import sqldf
from datetime import datetime


# Load environment variables
load_dotenv()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
MODEL_NAME = os.getenv("MODEL_NAME")
COMPANY_NAME = os.getenv("COMPANY_NAME")
WORKSHARE_FOLDER = os.getenv("WORKSHARE_FOLDER")
PROPOSAL_TEMPLATE = os.getenv("PROPOSAL_TEMPLATE")

# Initialize OpenAI client
client = OpenAI(
    base_url=OPENAI_BASE_URL,
    api_key=OPENAI_API_KEY
)

# Tool function to copy files
def copy_files(source_path: str, destination_path: str, file_pattern: str = "*") -> dict[str, str]:
    """
    Copies files from a source folder to a destination folder. Supports single files or multiple files via a pattern.
    
    Args:
        source_path (str): Path to the source folder or file (e.g., 'C:/Users/admin/source').
        destination_path (str): Path to the destination folder (e.g., 'C:/Users/admin/dest').
        file_pattern (str): File name or pattern (e.g., 'file.txt', '*.txt'). Defaults to '*' (all files).
    
    Returns:
        Dict[str, str]: Status ('success' or 'error') and message describing the outcome.
    
    Example:
        copy_files('C:/Users/admin/source', 'C:/Users/admin/dest', 'file.txt')
        copy_files('C:/Users/admin/source', 'C:/Users/admin/dest', '*.txt')
    """
    try:
        # Normalize paths
        source_path = os.path.normpath(source_path)
        destination_path = os.path.normpath(destination_path)

        # Validate source path
        if not os.path.exists(source_path):
            return {"status": "error", "message": f"Source path does not exist: {source_path}"}

        # Create destination folder if it doesn't exist
        os.makedirs(destination_path, exist_ok=True)

        # Handle single file
        if os.path.isfile(source_path):
            dest_file = os.path.join(destination_path, os.path.basename(source_path))
            shutil.copy2(source_path, dest_file)
            return {"status": "success", "message": f"Copied {os.path.basename(source_path)} to {destination_path}"}

        # Handle multiple files with pattern
        source_files = glob.glob(os.path.join(source_path, file_pattern))
        if not source_files:
            return {"status": "error", "message": f"No files found matching '{file_pattern}' in {source_path}"}

        copied_files = []
        for src_file in source_files:
            if os.path.isfile(src_file):
                dest_file = os.path.join(destination_path, os.path.basename(src_file))
                shutil.copy2(src_file, dest_file)
                copied_files.append(os.path.basename(src_file))

        return {
            "status": "success",
            "message": f"Copied {len(copied_files)} file(s): {', '.join(copied_files)} to {destination_path}"
        }

    except PermissionError:
        return {"status": "error", "message": "Permission denied. Check folder access rights."}
    except shutil.SameFileError:
        return {"status": "error", "message": "Source and destination are the same file."}
    except Exception as e:
        return {"status": "error", "message": f"Failed to copy files: {str(e)}"}

# Excel file path
EXCEL_FILE = "opportunities.xlsx"
SHEET_NAME = "Opportunities"

def add_opportunity(customer_name: str, opp_name: str, opp_id: str = "", deal_size: str = "", stage: str = "", details: str = "",
                        submission_date: str = "", tender_briefing_date: str = "", review1_date: str = "",
                        review2_date: str = "", am_name: str = "", offshore: str = "", bcc_review_date: str = ""):
    """Add a new opportunity to the Excel file with an auto-incremented running number."""
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        next_no = df["no"].max() + 1 if not df.empty else 1
        if opp_id.lower() in df["opp_id"].str.lower().values:
                return f"Opportunity ID {opp_id} already exists."
        if opp_name.lower() in df["opp_name"].str.lower().values:
                return f"Opportunity name {opp_name} already exists."
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "no", "timestamp", "customer_name", "opp_id", "opp_name", "submission_date",
            "tender_briefing_date", "review1_date", "review2_date", "am_name", "offshore",
            "bcc_review_date", "deal_size", "stage", "details"
        ])
        next_no = 1

    new_opportunity = {
        "no": next_no,
        "customer_name": customer_name,
        "opp_id": opp_id,
        "opp_name": opp_name,
        "submission_date": submission_date,
        "tender_briefing_date": tender_briefing_date,
        "review1_date": review1_date,
        "review2_date": review2_date,
        "am_name": am_name,
        "offshore": offshore,
        "bcc_review_date": bcc_review_date,
        "deal_size": deal_size,
        "stage": stage,
        "details": details
    }
    df = pd.concat([df, pd.DataFrame([new_opportunity])], ignore_index=True)
    df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
    return f"Opportunity {opp_id} added for {customer_name}."

def update_opportunity(opp_id: str = "", opp_name: str = "", new_opp_id: str = "", customer_name: str = "",
                           submission_date: str = "", tender_briefing_date: str = "", review1_date: str = "",
                           review2_date: str = "", am_name: str = "", offshore: str = "", bcc_review_date: str = "",
                           deal_size: str = "", stage: str = "", details: str = ""):
    if not opp_id and not opp_name:
        return "Error: Either opp_id or opp_name must be provided."

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        return f"No opportunities found to update."

    mask = (
        (df["opp_id"].str.lower() == opp_id.lower()) if opp_id else False |
        (df["opp_name"].str.lower() == opp_name.lower()) if opp_name else False
    )
    if not df[mask].empty:
        if new_opp_id:
                df.loc[mask, "opp_id"] = new_opp_id
        if customer_name:
                df.loc[mask, "customer_name"] = customer_name
        if submission_date:
                df.loc[mask, "submission_date"] = submission_date
        if tender_briefing_date:
                df.loc[mask, "tender_briefing_date"] = tender_briefing_date
        if review1_date:
                df.loc[mask, "review1_date"] = review1_date
        if review2_date:
                df.loc[mask, "review2_date"] = review2_date
        if am_name:
                df.loc[mask, "am_name"] = am_name
        if offshore:
                df.loc[mask, "offshore"] = offshore
        if bcc_review_date:
                df.loc[mask, "bcc_review_date"] = bcc_review_date
        if deal_size:
                df.loc[mask, "deal_size"] = deal_size
        if stage:
                df.loc[mask, "stage"] = stage
        if details:
                df.loc[mask, "details"] = df.loc[mask, "details"]+ "\n" + details
        df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
        identifier = opp_id or opp_name
        return f"Opportunity {identifier} updated."
    else:
        identifier = opp_id or opp_name
        return f"No opportunity found for {identifier}."

def query_opportunities(sql_query: str = "SELECT * FROM opportunities LIMIT 4"):
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        return ["No opportunities found."]

    try:
        # Execute SQL query using pandasql
        result_df = sqldf(sql_query, {"opportunities": df})
        if result_df.empty:
            return ["No opportunities match the query."]
        
        # Format results as strings
        results = []
        for _, row in result_df.iterrows():
            row_str = "\n".join([f"{col}: {row[col]}" for col in result_df.columns])
            results.append(row_str)
        return results
    except Exception as e:
        return [f"SQL query error: {str(e)}"]

# Function schema for OpenAI function calling
copy_files_schema = {
    "name": "copy_files",
    "description": "Copies files from a source folder to a destination folder. Supports single files or patterns.",
    "parameters": {
        "type": "object",
        "properties": {
            "source_path": {
                "type": "string",
                "description": "Path to the source folder or file (e.g., 'C:/Users/admin/source')."
            },
            "destination_path": {
                "type": "string",
                "description": "Path to the destination folder (e.g., 'C:/Users/admin/dest')."
            },
            "file_pattern": {
                "type": "string",
                "description": "File name or pattern (e.g., 'file.txt', '*.txt'). Defaults to all files.",
                "default": "*"
            }
        },
        "required": ["source_path", "destination_path"]
    }
}

add_opportunity_schema = {
	"name": "add_opportunity",
	"description": "Add a new sales opportunity to the Excel file with a unique running number. Fails if opp_name already exists.",
	"parameters": {
        "type": "object",
        "properties": {
            "customer_name": {"type": "string", "description": "Client name (e.g., SingTel)"},
            "opp_id": {"type": "string", "description": "Opportunity ID from other system (optional)", "default": ""},
            "opp_name": {"type": "string", "description": "Unique opportunity name"},
            "submission_date": {"type": "string", "description": "Submission date (YYYY-MM-DD, optional)", "default": ""},
            "tender_briefing_date": {"type": "string", "description": "Tender briefing date (YYYY-MM-DD, optional)", "default": ""},
            "review1_date": {"type": "string", "description": "1st review date with offshore team (YYYY-MM-DD, optional)", "default": ""},
            "review2_date": {"type": "string", "description": "2nd review date with offshore team (YYYY-MM-DD, optional)", "default": ""},
            "am_name": {"type": "string", "description": "Account manager name (optional)", "default": ""},
            "offshore": {"type": "string", "description": "Offshore team delivering the project (optional)", "default": ""},
            "bcc_review_date": {"type": "string", "description": "Review date with manager (YYYY-MM-DD, optional)", "default": ""},
            "deal_size": {"type": "string", "description": "Deal size (e.g., 500k)"},
            "stage": {"type": "string", "description": "Sales stage (e.g., Proposal)"},
            "details": {"type": "string", "description": "Additional details or notes"}
        },
        "required": ["customer_name", "opp_name", "deal_size", "stage", "details"]
	}
}

update_opportunity_schema = {
    "name": "update_opportunity",
    "description": "Update an existing sales opportunity in the Excel file, identified by either opp_id or opp_name, including updating opp_id if provided.",
    "parameters": {
        "type": "object",
        "properties": {
            "opp_id": {"type": "string", "description": "Opportunity ID to identify or update the opportunity (optional)", "default": ""},
            "opp_name": {"type": "string", "description": "Opportunity name to identify the opportunity (optional)", "default": ""},
            "new_opp_id": {"type": "string", "description": "New opportunity ID to set (optional)", "default": ""},
            "customer_name": {"type": "string", "description": "Updated client name (optional)", "default": ""},
            "submission_date": {"type": "string", "description": "Updated submission date (YYYY-MM-DD, optional)", "default": ""},
            "tender_briefing_date": {"type": "string", "description": "Updated tender briefing date (YYYY-MM-DD, optional)", "default": ""},
            "review1_date": {"type": "string", "description": "Updated 1st review date (YYYY-MM-DD, optional)", "default": ""},
            "review2_date": {"type": "string", "description": "Updated 2nd review date (YYYY-MM-DD, optional)", "default": ""},
            "am_name": {"type": "string", "description": "Updated account manager name (optional)", "default": ""},
            "offshore": {"type": "string", "description": "Updated offshore team (optional)", "default": ""},
            "bcc_review_date": {"type": "string", "description": "Updated manager review date (YYYY-MM-DD, optional)", "default": ""},
            "deal_size": {"type": "string", "description": "Updated deal size (optional)", "default": ""},
            "stage": {"type": "string", "description": "Updated sales stage (optional)", "default": ""},
            "details": {"type": "string", "description": "Updated details (optional)", "default": ""}
        },
        "required": []
    }
}

query_opportunities_schema = {
    "name": "query_opportunities",
    "description": "Retrieve opportunities from the Excel file using an SQL query. Default: SELECT * FROM opportunities LIMIT 4.",
    "parameters": {
        "type": "object",
        "properties": {
            "sql_query": {"type": "string", "description": "SQL query to filter opportunities (e.g., SELECT * FROM opportunities WHERE opp_name = 'AI Platform')", "default": "SELECT * FROM opportunities LIMIT 4"}
        },
        "required": []
    }
}

# ----------------------------
# System Prompt for the Agent
# ----------------------------
SYSTEM_PROMPT = (
    f"You are a presales assistant at {COMPANY_NAME}, support sales opportunities. Current time is {datetime.now()}. Your tools are:\n"
    "1) 'lookup_patient_data' Copies files from a source folder to a destination folder.\n"
    f"-The proposal template is '{PROPOSAL_TEMPLATE}' in folder '{WORKSHARE_FOLDER}\\00 Latest Templates\\Proposal Template\\01 Development Proposal'.\n"
    f"-Opportunity folder is opportunity name, under customer name folder in '{WORKSHARE_FOLDER}'.\n"
    "i.e., opportunity 'Build AI Agents' for customer NTU has folder name '{WORKSHARE_FOLDER}\\NTU\\Build AI Agents'.\n"
    "2) 'add_opportunity' Add a new opportunity (unique opp_name, optional opp_id).\n"
    "3) 'update_opportunity' Update an opportunity by opp_id or opp_name, including setting new_opp_id or other fields..\n"
	"4) 'query_opportunities' Query opportunities using SQL.\n"
    # "When asked to append information to the `details` field of an opportunity:.\n"
    # "1) Identify the opportunity by `opp_id` or `opp_name` from the user’s request.\n"
    # "2) Use `query_opportunities` with an SQL query (e.g., `SELECT details FROM opportunities WHERE opp_name = 'AI Platform'`) to retrieve the existing `details`.\n"
    # "3) If no opportunity is found, respond: 'No opportunity found for [opp_id or opp_name].'.\n"
    # "4) Append the new information to the existing `details` (e.g., add a new line with the new text).\n"
    # "5) Use `update_opportunity` to update the opportunity with the modified `details`, keeping other fields unchanged.\n"
    # "6) Respond with: \"Details updated for [opp_id or opp_name]\".\n"
    # "Focus on sales opportunities and avoid unrelated topics. Use tools for all operations and provide concise, relevant responses.\n"
)

@cl.on_chat_start
async def start():
    await cl.Message(content=f"Welcome to Anh's Agent! I can help manage opportunity. Current time is {datetime.now()}.'").send()

@cl.on_message
async def main(message: cl.Message):
    # Create chat completion with function calling
    response = client.chat.completions.create(
        model=MODEL_NAME,  # Adjust model based on your endpoint
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": message.content}
        ],
        tools=[{"type": "function", "function": copy_files_schema},
               {"type": "function", "function": add_opportunity_schema},
               {"type": "function", "function": update_opportunity_schema},
               {"type": "function", "function": query_opportunities_schema}],
        tool_choice="auto"
    )

    # Process response
    message_response = response.choices[0].message
    if message_response.tool_calls:
        for tool_call in message_response.tool_calls:
            result_content = ''
            if tool_call.function.name == "copy_files":
                # Parse function arguments
                args = json.loads(tool_call.function.arguments)
                
                # Call the copy_files function
                result = copy_files(
                    args["source_path"],
                    args["destination_path"],
                    args.get("file_pattern", "*")
                )
                
                # Send result to Chainlit UI
                await cl.Message(content=result["message"]).send()
                
                # Append result to conversation for context
                result_content = json.dumps(result)
				
            elif tool_call.function.name == "add_opportunity":
                arguments = json.loads(tool_call.function.arguments)
                result = add_opportunity(
										customer_name=arguments["customer_name"],
                    opp_name=arguments["opp_name"],
										opp_id=arguments.get("opp_id", ""),
										deal_size=arguments.get("deal_size", ""),
										stage=arguments.get("stage", ""),
										details=arguments.get("details", ""),
										submission_date=arguments.get("submission_date", ""),
										tender_briefing_date=arguments.get("tender_briefing_date", ""),
										review1_date=arguments.get("review1_date", ""),
										review2_date=arguments.get("review2_date", ""),
										am_name=arguments.get("am_name", ""),
										offshore=arguments.get("offshore", ""),
										bcc_review_date=arguments.get("bcc_review_date", "")
								)
                # Send result to Chainlit UI
                await cl.Message(content=result).send()
                
				# Append result to conversation for context
                result_content = json.dumps(result)
                	
 
            elif tool_call.function.name == "update_opportunity":
                arguments = json.loads(tool_call.function.arguments)
                result = update_opportunity(
                    opp_id=arguments.get("opp_id", ""),
                    opp_name=arguments.get("opp_name", ""),
                    new_opp_id=arguments.get("new_opp_id", ""),
                    customer_name=arguments.get("customer_name", ""),
                    submission_date=arguments.get("submission_date", ""),
                    tender_briefing_date=arguments.get("tender_briefing_date", ""),
                    review1_date=arguments.get("review1_date", ""),
                    review2_date=arguments.get("review2_date", ""),
                    am_name=arguments.get("am_name", ""),
                    offshore=arguments.get("offshore", ""),
                    bcc_review_date=arguments.get("bcc_review_date", ""),
                    deal_size=arguments.get("deal_size", ""),
                    stage=arguments.get("stage", ""),
                    details=arguments.get("details", "")
                )  
                # Send result to Chainlit UI
                await cl.Message(content=result).send()
                
				# Append result to conversation for context
                result_content = json.dumps(result)
            
            elif tool_call.function.name == "query_opportunities":
                arguments = json.loads(tool_call.function.arguments)
                sql_query=arguments.get("sql_query", "SELECT * FROM opportunities LIMIT 4")
               
                await cl.Message(content=sql_query).send()
                result = query_opportunities(sql_query=sql_query)

				# Append result to conversation for context
                result_content = "\n".join(result)
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {"role": "system", "content": f"Current time is {datetime.now()}. You are a presales assistant at {COMPANY_NAME}. Use the tool results to provide concise, relevant responses. Focus on sales opportunities."},
                    {"role": "user", "content": message.content},
                    {"role": "assistant", "content": None, "tool_calls": message_response.tool_calls},
                    {
                        "role": "tool",
                        "content": result_content,
                        "tool_call_id": tool_call.id
                    }
                ]
            )  
            await cl.Message(content=response.choices[0].message.content).send()	 					
                
    else:
        await cl.Message(content=message_response.content).send()