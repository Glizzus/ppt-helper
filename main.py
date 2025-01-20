import sys
import ollama
from ppt import Schema, create_ppt

ollama_client = ollama.Client()

prompt = """Generate the colors for a PowerPoint presentation as JSON.
This includes the background color and the theme colors.

Apply UI/UX principles to choose the colors. Choose colors reminiscent of the
desired mood or theme of the presentation. A client will ask you to create a
template for their presentation, and you need to generate the colors for it.

Client requirements:
The presentation will be about
- VS Code
- Docker
- Dev Containers
- Software Engineering
"""

full_str_response = ''
stream = ollama_client.generate("phi3", prompt, format=Schema.model_json_schema(), stream=True)

for chunk in stream:
    sys.stdout.write(chunk.response)
    full_str_response += chunk.response

response_schema = Schema.model_validate_json(full_str_response)

create_ppt(response_schema, "C:/Users/Glizzus/Documents/dev_container_template.pptx")
