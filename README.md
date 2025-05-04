### Table of Contents

1. [Installation](#installation)
2. [Project Motivation](#motivation)
3. [File Descriptions](#files)
4. [Results](#results)
5. [Licensing, Authors, and Acknowledgements](#licensing)

## Installation <a name="installation"></a>

This is the Python app that can be run with
1. Python versions 3.*
2. chainlit library to easy create chat UI
3. pandasql library to easy query data using SQL
- If you run 'presale_agent.py', you should install openai package
- If you run 'presale_agent_langchain.py' you should install `pip install langchain langchain-community langchain-openai chainlit`.<br>

#### example run: 
- `chainlit run presale_agent.py --port 8000`
- `chainlit run presale_agent_langchain.py --port 8001`

## Project Motivation<a name="motivation"></a>

As a presale I need to keep track the opportunity that I am assigned to support. The agent can help me:
- Copy the files between folders. Copy proposal template to new opportunity folder
- Create a new opportunity
- Update an opportunity with more details and notes
- Query opportunities, like the urgent opportunity that I need submit proposal in 3 days, the big deal size opportunities, the opportunities that I have to take actions today
- Draw chart of opportunities or other data if provided

## File Descriptions <a name="files"></a>

#### presale_agent.py
The file contain scripts of agent using OpenAI

#### presale_agent_langchain.py
The file contain scripts of agent using Langchain

## Results<a name="results"></a>
- The agent help me talk to my data of opportunities in nature language
- I tried to use Mem0 to enhance the memory of the agent but the limit is 1000 query per month only.
- I tried OpenAI’s tool-chaining capabilities but it does not work. Seem OpenAI does not support it.<br>
The prompt of tool-chaining that does not work with OpenAI. Only `query_opportunities` is called:<br>
>"When asked to append information to the `details` field (e.g., 'Append [text] to details of [opp_name or opp_id]'):\n"<br>
>"1. Extract the opp_id or opp_name and the text to append.\n"<br>
>"2. Call `query_opportunities` with an SQL query (e.g., `SELECT details, opp_id, opp_name FROM opportunities WHERE opp_name = '[opp_name]' OR opp_id = '[opp_id]'`) to retrieve `details` and identifiers.\n"<br>
>"3. If no opportunity is found, respond: 'No opportunity found for [opp_id or opp_name].'\n"<br>
>"4. Append the new text to the existing `details` with a newline (e.g., existing_details + '\\n' + new_text). If `details` is empty, use the new text.\n"<br>
>"5. Call `update_opportunity` in the SAME RESPONSE, using the same opp_id or opp_name, setting `details` to the appended text. Do not modify other fields.\n"<br>
>"6. Respond with: 'Details updated for [opp_id or opp_name].'"<br>

- The above prompt work well with Langchain agent

## Licensing, Authors, Acknowledgements<a name="licensing"></a>
Feel free to use the code here as you would like! <br>
