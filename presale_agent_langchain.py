import os
import pandas as pd
from datetime import datetime, UTC
from langchain.agents import create_openai_tools_agent, AgentExecutor
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_openai import ChatOpenAI
from langchain.tools import tool
from pandasql import sqldf
import shutil
import chainlit as cl
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

MODEL_NAME = os.getenv("MODEL_NAME")
COMPANY_NAME = os.getenv("COMPANY_NAME")
WORKSHARE_FOLDER = os.getenv("WORKSHARE_FOLDER")
PROPOSAL_TEMPLATE = os.getenv("PROPOSAL_TEMPLATE")

# Initialize Grok 3 via xAI API
try:
    llm = ChatOpenAI(
        openai_api_base=OPENAI_BASE_URL,
        openai_api_key=OPENAI_API_KEY,
        model_name=MODEL_NAME,
        temperature=0.7,
        max_tokens=500
    )
    if not OPENAI_API_KEY:
        raise ValueError("OPENAI_API_KEY environment variable is not set.")
except Exception as e:
    raise Exception(f"Failed to initialize Grok 3: {str(e)}")

# Excel file path
EXCEL_FILE = "opportunities.xlsx"
SHEET_NAME = "Opportunities"

# Initialize Excel file
def init_excel():
    columns = [
        "no", "timestamp", "customer_name", "opp_id", "opp_name", "submission_date",
        "tender_briefing_date", "review1_date", "review2_date", "am_name", "offshore",
        "bcc_review_date", "deal_size", "stage", "details"
    ]
    if not os.path.exists(EXCEL_FILE):
        pd.DataFrame(columns=columns).to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)

init_excel()

# Define tools
@tool
def copy_files_or_folder(source_path: str, destination_path: str) -> str:
    """
    Copy a file or folder from source_path to destination_path.
    
    Args:
        source_path (str): Path to the source file or folder.
        destination_path (str): Path to the destination.
    
    Returns:
        str: Success or error message.
    """
    try:
        if not os.path.exists(source_path):
            return f"Error: Source path {source_path} does not exist."
        if os.path.isfile(source_path):
            shutil.copy2(source_path, destination_path)
            return f"File copied from {source_path} to {destination_path}."
        elif os.path.isdir(source_path):
            shutil.copytree(source_path, destination_path, dirs_exist_ok=True)
            return f"Folder copied from {source_path} to {destination_path}."
        else:
            return f"Error: {source_path} is neither a file nor a folder."
    except Exception as e:
        return f"Error copying: {str(e)}"

@tool
def add_opportunity(customer_name: str, opp_name: str, deal_size: str, stage: str, details: str,
                    opp_id: str = "", submission_date: str = "", tender_briefing_date: str = "",
                    review1_date: str = "", review2_date: str = "", am_name: str = "",
                    offshore: str = "", bcc_review_date: str = "") -> str:
    """
    Add a new sales opportunity to the Excel file using pandas. Fails if opp_name already exists.
    
    Args:
        customer_name (str): Client name (e.g., SingTel).
        opp_name (str): Unique opportunity name.
        deal_size (str): Deal size (e.g., 500k).
        stage (str): Sales stage (e.g., Proposal).
        details (str): Additional details or notes.
        opp_id (str, optional): Opportunity ID from other system.
        submission_date (str, optional): Submission date (YYYY-MM-DD).
        tender_briefing_date (str, optional): Tender briefing date (YYYY-MM-DD).
        review1_date (str, optional): 1st review date (YYYY-MM-DD).
        review2_date (str, optional): 2nd review date (YYYY-MM-DD).
        am_name (str, optional): Account manager name.
        offshore (str, optional): Offshore team.
        bcc_review_date (str, optional): Manager review date (YYYY-MM-DD).
    
    Returns:
        str: Success or error message.
    """
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        df = pd.DataFrame(columns=[
            "no", "timestamp", "customer_name", "opp_id", "opp_name", "submission_date",
            "tender_briefing_date", "review1_date", "review2_date", "am_name", "offshore",
            "bcc_review_date", "deal_size", "stage", "details"
        ])

    # Check for duplicate opp_name
    if not df[df["opp_name"].str.lower() == opp_name.lower()].empty:
        return f"Opportunity name {opp_name} already exists."

    # Get next no
    next_no = df["no"].max() + 1 if not df.empty else 1

    # Prepare new opportunity
    timestamp = datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S")
    new_row = {
        "no": next_no, "timestamp": timestamp, "customer_name": customer_name, "opp_id": opp_id,
        "opp_name": opp_name, "submission_date": submission_date, "tender_briefing_date": tender_briefing_date,
        "review1_date": review1_date, "review2_date": review2_date, "am_name": am_name, "offshore": offshore,
        "bcc_review_date": bcc_review_date, "deal_size": deal_size, "stage": stage, "details": details
    }
    try:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
        return f"Opportunity {opp_name} added for {customer_name}."
    except Exception as e:
        return f"Error adding opportunity: {str(e)}"

@tool
def query_opportunities(sql_query: str = "SELECT * FROM opportunities LIMIT 3") -> str:
    """
    Retrieve opportunities from the Excel file using an SQL query.
    
    Args:
        sql_query (str): SQL query to filter opportunities (default: SELECT * FROM opportunities LIMIT 3).
    
    Returns:
        str: Query results or error message.
    """
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        return "No opportunities found."

    try:
        result_df = sqldf(sql_query, {"opportunities": df})
        if result_df.empty:
            return "No opportunities match the query."
        
        results = []
        for _, row in result_df.iterrows():
            row_str = "\n".join([f"{col}: {row[col]}" for col in result_df.columns])
            results.append(row_str)
        return "\n\n".join(results)
    except Exception as e:
        return f"SQL query error: {str(e)}"

@tool
def update_opportunity(opp_id: str = "", opp_name: str = "", new_opp_id: str = "", customer_name: str = "",
                       submission_date: str = "", tender_briefing_date: str = "", review1_date: str = "",
                       review2_date: str = "", am_name: str = "", offshore: str = "", bcc_review_date: str = "",
                       deal_size: str = "", stage: str = "", details: str = "") -> str:
    """
    Update an existing sales opportunity in the Excel file using pandas, identified by either opp_id or opp_name.
    
    Args:
        opp_id (str, optional): Opportunity ID to identify the opportunity.
        opp_name (str, optional): Opportunity name to identify the opportunity.
        new_opp_id (str, optional): New opportunity ID to set.
        customer_name (str, optional): Updated client name.
        submission_date (str, optional): Updated submission date (YYYY-MM-DD).
        tender_briefing_date (str, optional): Updated tender briefing date (YYYY-MM-DD).
        review1_date (str, optional): Updated 1st review date (YYYY-MM-DD).
        review2_date (str, optional): Updated 2nd review date (YYYY-MM-DD).
        am_name (str, optional): Updated account manager name.
        offshore (str, optional): Updated offshore team.
        bcc_review_date (str, optional): Updated manager review date (YYYY-MM-DD).
        deal_size (str, optional): Updated deal size.
        stage (str, optional): Updated sales stage.
        details (str, optional): Updated details.
    
    Returns:
        str: Success or error message.
    """
    if not opp_id and not opp_name:
        return "Error: Either opp_id or opp_name must be provided."

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except FileNotFoundError:
        return "No opportunities found to update."

    # Check if opportunity exists
    mask = (
        (df["opp_id"].str.lower() == opp_id.lower()) if opp_id else False |
        (df["opp_name"].str.lower() == opp_name.lower()) if opp_name else False
    )
    if not df[mask].any().any():
        identifier = opp_id or opp_name
        return f"No opportunity found for {identifier}."

    try:
        # Apply updates
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
            df.loc[mask, "details"] = details
        df.loc[mask, "timestamp"] = datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S")
        df.to_excel(EXCEL_FILE, sheet_name=SHEET_NAME, index=False)
        identifier = opp_id or opp_name
        return f"Opportunity {identifier} updated."
    except Exception as e:
        return f"Error updating opportunity: {str(e)}"

# Define prompt
prompt = ChatPromptTemplate.from_messages([
    ("system", 
     f"""
     You are a presales assistant at {COMPANY_NAME}, managing sales opportunities in an Excel file and handling file operations. Your tools are:
     - `copy_files_or_folder`: Copy a file or folder from source_path to destination_path.
     - `add_opportunity`: Add a new opportunity to the Excel file (unique opp_name, optional opp_id).
     - `query_opportunities`: Query opportunities using SQL (table: opportunities, columns: no, timestamp, customer_name, opp_id, opp_name, submission_date, tender_briefing_date, review1_date, review2_date, am_name, offshore, bcc_review_date, deal_size, stage, details). Default: SELECT * FROM opportunities LIMIT 3.
     - `update_opportunity`: Update an opportunity by opp_id or opp_name, including setting new_opp_id or other fields.
     
     When used tool `copy_files_or_folder`:
     - If metioned to Development Proposal Template, it is the file at path '{WORKSHARE_FOLDER}\\00 Latest Templates\\Proposal Template\\01 Development Proposal\\{PROPOSAL_TEMPLATE}'.
     - If metioned to opportunity folder, the folder name is opportunity name, under customer name folder in '{WORKSHARE_FOLDER}'. i.e., opportunity 'Build AI Agents' for customer NTU has folder's path '{WORKSHARE_FOLDER}\\NTU\\Build AI Agents'.\n"

     When asked to append information to the `details` field (e.g., 'Append [text] to details of [opp_name or opp_id]'):
     1. Extract the opp_id or opp_name and the text to append from the user's request.
     2. Call `query_opportunities` with an SQL query (e.g., `SELECT details, opp_id, opp_name FROM opportunities WHERE opp_name = '[opp_name]' OR opp_id = '[opp_id]'`) to retrieve the existing `details` and identifiers.
     3. If no opportunity is found, respond: 'No opportunity found for [opp_id or opp_name].'
     4. Append the new information to the existing `details` with a newline (e.g., existing_details + '\\n' + new_text). If `details` is empty, use the new text directly.
     5. Call `update_opportunity` with the same opp_id or opp_name used in the query, setting the `details` field to the appended text. Do not modify other fields.
     6. Respond with: 'Details updated for [opp_id or opp_name].'

     For other requests (e.g., querying, adding opportunities, or copying files), use the appropriate tool and respond concisely. Focus on sales opportunities (e.g., for SingTel, U Mobile) and file operations. Always use tools for data and file operations. Do not generate tool calls manually; the agent will handle tool execution.
     """
    ),
    MessagesPlaceholder(variable_name="chat_history"),
    ("human", "{input}"),
    MessagesPlaceholder(variable_name="agent_scratchpad")
])

# Initialize agent
tools = [copy_files_or_folder, add_opportunity, query_opportunities, update_opportunity]
agent = create_openai_tools_agent(llm, tools, prompt)
executor = AgentExecutor(agent=agent, tools=tools, verbose=True)

# Chainlit event handlers
@cl.on_chat_start
async def on_chat_start():
    cl.user_session.set("user_id", "presales_01")
    cl.user_session.set("chat_history", [])
    await cl.Message(content="Hi! How can I assist with your sales opportunities or file operations today?").send()

@cl.on_message
async def on_message(message: cl.Message):
    user_id = cl.user_session.get("user_id")
    chat_history = cl.user_session.get("chat_history")

    # Execute agent
    try:
        response = await cl.make_async(executor.invoke)({
            "input": message.content,
            "chat_history": chat_history
        })

        # Update chat history
        chat_history.append(("human", message.content))
        chat_history.append(("assistant", response["output"]))
        cl.user_session.set("chat_history", chat_history)

        # Send response
        await cl.Message(content=response["output"]).send()

    except Exception as e:
        error_message = f"Error: Failed to process request. {str(e)}"
        await cl.Message(content=error_message).send()