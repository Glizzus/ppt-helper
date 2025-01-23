from pathlib import Path
import sys
from typing import TextIO
import ollama
import argparse
from ppt import Schema, create_ppt


parser = argparse.ArgumentParser(
    description="Generate a template for a PowerPoint presentation"
)
parser.add_argument(
    'dir',
    type=Path,
    nargs='?',
    default=Path.cwd(),
)

args = parser.parse_args()

ollama_client = ollama.Client()


def read_until_empty_line(input_stream: TextIO = sys.stdin) -> str:
    lines: list[str] = []
    for line in input_stream:
        if line.strip() == "":
            break
        lines.append(line)
    return "".join(lines)


system_prompt = """You are tasked with generating a JSON object
that recommends colors and typefaces for a PowerPoint presentation.
Use sound UI/UX and typography principles. Provide a brief reason for EVERY choice.

Requirements for your JSON response:
1. All color codes must be 6-digit uppercase hex values (e.g., "#FFFFFF").
2. Include one background color and a set of theme colors.
3. Include a header font and a body typeface.
4. Return valid JSON"""

user_prompt = read_until_empty_line(sys.stdin)

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

create_ppt(response_schema, args.dir / "output.pptx")
