import sys
import ollama
from ppt import Schema, create_ppt


ollama_client = ollama.Client()

system_prompt = """You are tasked with generating a JSON object
that recommends colors and typefaces for a PowerPoint presentation.
Use sound UI/UX and typography principles. Provide a brief reason for EVERY choice.

Requirements for your JSON response:
1. All color codes must be 6-digit uppercase hex values (e.g., "#FFFFFF").
2. Include one background color and a set of theme colors.
3. Include a header font and a body typeface.
4. Return valid JSON"""

user_prompt = """Our client requires a presentation on:
- ElasticSearch
- Logging
- Observability

They want a clean, modern look that conveys a technical, professional mood. 
Please generate a color scheme and font recommendations that fit these requirements."""

full_str_response = ""
stream = ollama_client.generate(
    model="llama3.2",
    system=system_prompt,
    prompt=user_prompt,
    format=Schema.model_json_schema(),
    stream=True,
    options={
        "temperature": 0
    },
)

for chunk in stream:
    sys.stdout.write(chunk.response)
    full_str_response += chunk.response

response_schema = Schema.model_validate_json(full_str_response)

create_ppt(response_schema, "C:/Users/Glizzus/Documents/dev_container_template.pptx")
