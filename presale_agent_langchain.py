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
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import base64
from io import BytesIO
from RestrictedPython import compile_restricted, safe_builtins
from RestrictedPython.Eval import default_guarded_getitem
import re
import tempfile
import uuid
import logging
from langchain.agents import Tool

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

MODEL_NAME = os.getenv("MODEL_NAME")
COMPANY_NAME = os.getenv("COMPANY_NAME")
WORKSHARE_FOLDER = os.getenv("WORKSHARE_FOLDER")
PROPOSAL_TEMPLATE = os.getenv("PROPOSAL_TEMPLATE")

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
    raise Exception(f"Failed to initialize LLM: {str(e)}")

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

# Function to clean LLM-generated code
def clean_code(code: str) -> str:
    """
    Remove Markdown code block markers and other non-Python content from LLM-generated code.
    
    Args:
        code (str): Raw code string from LLM.
    
    Returns:
        str: Cleaned Python code.
    """
    code = re.sub(r'^```(?:python)?\s*\n|\n```$', '', code, flags=re.MULTILINE)
    code = code.strip()
    lines = [line for line in code.split('\n') if line.strip() and not line.strip().startswith('#')]
    return '\n'.join(lines)


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

    if not df[df["opp_name"].str.lower() == opp_name.lower()].empty:
        return f"Opportunity name {opp_name} already exists."

    next_no = df["no"].max() + 1 if not df.empty else 1
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
        str: Query results as a string for display.
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

    mask = (
        (df["opp_id"].str.lower() == opp_id.lower()) if opp_id else False |
        (df["opp_name"].str.lower() == opp_name.lower()) if opp_name else False
    )
    if not df[mask].any().any():
        identifier = opp_id or opp_name
        return f"No opportunity found for {identifier}."

    try:
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

def draw_chart_tool(chart_query: str) -> str:
    """
    Generate a Matplotlib chart based on the user's query, either using direct data input or querying data from the opportunities Excel file.
    
    Args:
        chart_query (str): User query describing the chart. Supports:
            - Direct data (e.g., "Draw a pie chart orange 30, apple 25, cucumber 40").
            - Excel query (e.g., "Draw a bar chart of opportunities by stage where customer_name = 'SingTel'").
    
    Returns:
        str: Path to temporary image file or error message.
    """
    try:
        matplotlib.use('Agg')
        
        # Check for direct data input (e.g., "orange 30, apple 25, cucumber 40")
        direct_data_match = re.match(r'Draw a \w+ chart\s+([\w\s]+\d+(?:,\s*[\w\s]+\d+)*)\s*$', chart_query, re.IGNORECASE)
        if direct_data_match:
            data_str = direct_data_match.group(1)
            # Parse item-value pairs
            pairs = re.findall(r'([\w\s]+?)\s+(\d+)(?:,|$)', data_str.strip())
            if pairs:
                items, values = zip(*pairs)
                plot_df = pd.DataFrame({
                    'item': [item.strip() for item in items],
                    'value': [int(v) for v in values]
                })
            else:
                return "Error: Invalid direct data format. Use 'item value, item value, ...' (e.g., 'orange 30, apple 25')."
        else:
            # Fallback to Excel query
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            if df.empty:
                return "No opportunity data available to plot."

            # Handle "opp by [column]" queries
            by_match = re.match(r'Draw a \w+ chart\s+opp\s+by\s+(\w+)(?:\s+where\s+(.+))?$', chart_query, re.IGNORECASE)
            if by_match:
                group_column = by_match.group(1)
                where_clause = by_match.group(2) or ""
                sql_query = f"SELECT {group_column}, COUNT(*) as count FROM opportunities"
                if where_clause:
                    sql_query += f" WHERE {where_clause}"
                sql_query += f" GROUP BY {group_column}"
            else:
                sql_match = re.search(r'where\s+(.+?)(?:\s+|$)', chart_query, re.IGNORECASE)
                sql_query = f"SELECT * FROM opportunities" if not sql_match else f"SELECT * FROM opportunities WHERE {sql_match.group(1)}"
            
            try:
                plot_df = sqldf(sql_query, {"opportunities": df})
                if plot_df.empty:
                    return "No opportunities match the query."
            except Exception as e:
                return f"SQL query error: {str(e)}"

        # Determine columns for chart
        columns = ['item', 'value'] if 'item' in plot_df.columns else plot_df.columns.tolist()
        
        code_prompt = f"""
        Generate concise Python code to create a Matplotlib chart based on the user query: "{chart_query}".
        The data is in a pandas DataFrame 'df' with columns: {columns}.
        - Do NOT include import statements; use pre-defined variables: pd (pandas), plt (matplotlib.pyplot), np (numpy), BytesIO.
        - Do NOT use base64 encoding or generate base64 strings (e.g., no 'data:image/png;base64,...').
        - Create a chart (e.g., bar, line, pie, histogram) based on the query.
        - For direct data (columns 'item', 'value'), use 'item' for labels and 'value' for quantities.
        - For Excel data with 'count' column (e.g., grouped by stage), use the first column for labels and 'count' for quantities.
        - Set default size to (8, 5). Can be adjusted if needed.
        - Change chart colors if user requests.
        - Include appropriate title; x-label, y-label, and grid are optional.
        - Handle 'deal_size' in Excel data by converting strings like '500k' to numeric (e.g., df['deal_size'].str.replace('k', '000').astype(float)).
        - Save the chart to a BytesIO buffer as PNG with dpi=80, then write to a temporary file (provided as 'temp_file').
        - Return only the Python code as a string, without explanations, comments, or Markdown formatting (e.g., ```python).
        - Do not use plt.show(), eval(), exec(), or base64.
        - Example for direct data pie chart:
        plt.figure(figsize=(8, 5))
        plt.pie(df['value'], labels=df['item'], autopct='%1.1f%%')
        plt.title('Chart')
        buffer = BytesIO()
        plt.savefig(buffer, format='png', dpi=80)
        plt.close()
        with open(temp_file, 'wb') as f:
            f.write(buffer.getvalue())
        buffer.close()
        - Example for Excel grouped data pie chart:
        plt.figure(figsize=(8, 5))
        plt.pie(df['count'], labels=df['{columns[0] if columns[0] != 'count' else columns[1]}'], autopct='%1.1f%%')
        plt.title('Opportunities by {columns[0] if columns[0] != 'count' else columns[1]}')
        buffer = BytesIO()
        plt.savefig(buffer, format='png', dpi=80)
        plt.close()
        with open(temp_file, 'wb') as f:
            f.write(buffer.getvalue())
        buffer.close()
        """
        code_response = llm.invoke(code_prompt)
        plot_code = code_response.content.strip()
        logger.debug(f"Raw LLM-generated code: {plot_code}")

        cleaned_code = clean_code(plot_code)
        if not cleaned_code:
            return "Error: LLM generated empty or invalid code."

        if any(unsafe in cleaned_code for unsafe in ['eval(', 'exec(', 'import ', 'base64.', 'data:image/png;base64']):
            return f"Error: Generated code contains unsafe operations (eval/exec/import/base64): {cleaned_code[:100]}..."

        if not cleaned_code.startswith(('plt.', 'df.', 'counts =', 'sizes =')):
            return f"Error: Generated code appears invalid: {cleaned_code[:50]}..."

        temp_file = os.path.join(tempfile.gettempdir(), f"chart_{uuid.uuid4()}.png")
        restricted_globals = {
            '__builtins__': safe_builtins,
            'df': plot_df,
            'pd': pd,
            'plt': plt,
            'np': np,
            'BytesIO': BytesIO,
            '__getitem__': default_guarded_getitem,
            '_getitem_': default_guarded_getitem,
            'temp_file': temp_file,
            'open': open
        }

        try:
            byte_code = compile_restricted(cleaned_code, '<inline>', 'exec')
            exec(byte_code, restricted_globals)
            if os.path.exists(temp_file):
                file_size = os.path.getsize(temp_file) / 1024
                logger.debug(f"Chart saved to {temp_file}, size: {file_size:.2f} KB")
                return temp_file
            else:
                return f"Error: Failed to save chart to temporary file: {temp_file}. Code: {cleaned_code}"
        except Exception as e:
            error_msg = str(e)
            if 'Eval calls are not allowed' in error_msg:
                return f"Error: Generated code contains eval-like operations: {error_msg}. Code: {cleaned_code}"
            return f"Error executing chart code: {error_msg}. Code: {cleaned_code}"
    except Exception as e:
        return f"Error generating chart: {str(e)}"

draw_chart = Tool(
    name="draw_chart",
    func=draw_chart_tool,
    description="Generate a chart based on the user's query and return the file path.",
    return_direct=True  #This stops the agent after tool executes
)

# Define prompt
prompt = ChatPromptTemplate.from_messages([
    ("system", 
     f"""
     You are a presales assistant at {COMPANY_NAME}, managing sales opportunities in an Excel file and handling file operations. Your tools are:
     - `copy_files_or_folder`: Copy a file or folder from source_path to destination_path.
     - `add_opportunity`: Add a new opportunity to the Excel file (unique opp_name, optional opp_id).
     - `query_opportunities`: Query opportunities using SQL (table: opportunities, columns: no, timestamp, customer_name, opp_id, opp_name, submission_date, tender_briefing_date, review1_date, review2_date, am_name, offshore, bcc_review_date, deal_size, stage, details). Default: SELECT * FROM opportunities LIMIT 3.
     - `update_opportunity`: Update an opportunity by opp_id or opp_name, including setting new_opp_id or other fields.
     - `draw_chart`: Generate a Matplotlib chart based on a user query. Supports two modes:
         - Direct data input (e.g., "Draw a pie chart orange 30, apple 25, cucumber 40"), where items and numeric values are provided as comma-separated pairs.
         - Excel query (e.g., "Draw a bar chart of opportunities by stage" or "Draw a pie chart opp by stage where customer_name = 'SingTel'"), using data from the opportunities Excel file.

     When used tool `copy_files_or_folder`:
     - If mentioned to Development Proposal Template, it is the file at path '{WORKSHARE_FOLDER}\\00 Latest Templates\\Proposal Template\\01 Development Proposal\\{PROPOSAL_TEMPLATE}'.
     - If mentioned to opportunity folder, the folder name is opportunity name, under customer name folder in '{WORKSHARE_FOLDER}'. i.e., opportunity 'Build AI Agents' for customer NTU has folder's path '{WORKSHARE_FOLDER}\\NTU\\Build AI Agents'.

     When asked to append information to the `details` field (e.g., 'Append [text] to details of [opp_name or opp_id]'):
     1. Extract the opp_id or opp_name and the text to append from the user's request.
     2. Call `query_opportunities` with an SQL query (e.g., `SELECT details, opp_id, opp_name FROM opportunities WHERE opp_name = '[opp_name]' OR opp_id = '[opp_id]'`) to retrieve the existing `details` and identifiers.
     3. If no opportunity is found, respond: 'No opportunity found for [opp_id or opp_name].'
     4. Append the new information to the existing `details` with a newline (e.g., existing_details + '\\n' + new_text). If `details` is empty, use the new text directly.
     5. Call `update_opportunity` with the same opp_id or opp_name used in the query, setting the `details` field to the appended text. Do not modify other fields.
     6. Respond with: 'Details updated for [opp_id or opp_name].'

     When asked to draw a chart:
     1. Determine if the query provides direct data (e.g., "Draw a pie chart orange 30, apple 25, cucumber 40") or references Excel data (e.g., "Draw a pie chart opp by stage").
     2. Call `draw_chart` with the user query. The tool will:
        - For direct data: Parse item-value pairs (e.g., "orange 30, apple 25") into a DataFrame with 'item' and 'value' columns.
        - For Excel data: If query is "opp by [column]" (e.g., "opp by stage"), generate SQL like "SELECT [column], COUNT(*) FROM opportunities GROUP BY [column]"; otherwise, extract WHERE clause (e.g., "customer_name = 'SingTel'") or use "SELECT * FROM opportunities".
        - Generate Python code to create the chart using the data.
        - Save the chart as a PNG to a temporary file.
     3. The tool returns the path to the temporary file.
     4. Respond with the raw path to the temporary file (e.g., '/tmp/chart_1234567890.png'). UI will handle the file and display it.
     5. If the chart generation fails, respond with the error message (e.g., 'Error generating chart: [error message]').

     If you need to know the current date or time, use the date time now [{datetime.now(UTC).isoformat()}].
     
     For other requests (e.g., querying, adding opportunities, or copying files), use the appropriate tool and respond concisely. Focus on sales opportunities (e.g., for SingTel, U Mobile) and file operations. Always use tools for data and file operations. Do not generate tool calls manually; the agent will handle tool execution.
     """
    ),
    MessagesPlaceholder(variable_name="chat_history"),
    ("human", "{input}"),
    MessagesPlaceholder(variable_name="agent_scratchpad")
])

# Initialize agent
tools = [copy_files_or_folder, add_opportunity, query_opportunities, update_opportunity, draw_chart]
agent = create_openai_tools_agent(llm, tools, prompt)
executor = AgentExecutor(agent=agent, tools=tools, verbose=True)

# Chainlit event handlers
@cl.on_chat_start
async def on_chat_start():
    cl.user_session.set("user_id", "presales_01")
    cl.user_session.set("chat_history", [])
    await cl.Message(content="Hi! How can I assist with your sales opportunities, file operations, or charts today?").send()

@cl.on_message
async def on_message(message: cl.Message):
    user_id = cl.user_session.get("user_id")
    chat_history = cl.user_session.get("chat_history", [])

    try:
        response = await cl.make_async(executor.invoke)({
            "input": message.content,
            "chat_history": chat_history
        })

        chat_history.append(("human", message.content))
        chat_history.append(("assistant", response["output"]))
        cl.user_session.set("chat_history", chat_history)

        if os.path.exists(response["output"]):
            logger.debug(f"Sending image to UI: {response['output']}")
            # Create and send the image
            image = cl.Image(path=str(response["output"]), name="My Image")
            await cl.Message(content="Here's the chart:", elements=[image]).send()
            # Delay cleanup to ensure image is rendered
            try:
                logger.debug(f"Cleaning up temporary file: {response['output']}")
                os.remove(response["output"])
            except Exception as e:
                logger.warning(f"Failed to clean up temporary file {response['output']}: {str(e)}")
        else:
            await cl.Message(content=response["output"]).send()

    except Exception as e:
        error_message = f"Error: Failed to process request. {str(e)}"
        logger.error(error_message)
        await cl.Message(content=error_message).send()