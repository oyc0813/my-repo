//Install the OpenAI Python SDK and execute the code below to generate a haiku for free using the gpt-4o-mini model.
pip install openai

from openai import OpenAI

client = OpenAI(
  api_key="xxx"
)

completion = client.chat.completions.create(
  model="gpt-4o-mini",
  store=True,
  messages=[
    {"role": "user", "content": "write a haiku about ai"}
  ]
)

print(completion.choices[0].message);
