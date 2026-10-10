# AI_Resume_Generator
Web-based UI that uses AI to generate a JSON containing all the pieces of resume information, then parses that information into a well-formatted docx document.

It is currently configured to a locally hosted LLM with the IP address hard coded into main.py
The app is run using uvicorn:
`uvicorn backend.main:app --reload`
Required python packages are fastapi, pydantic, and python-docx
