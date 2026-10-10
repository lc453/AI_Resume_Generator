from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# Document creation and saving
from docx import Document
import resume_interpreter
from io import BytesIO
from fastapi.responses import StreamingResponse

# Setup for the LLM connection
import requests
import re
import json
BASE_URL = "http://192.168.68.154:5001"

BASE_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"
response = requests.get(f"{BASE_URL}/v1/models")
modelname = response.json()["data"][0]["id"]

JOBINFO_SYSTEM_PROMPT = """
You are an expert Career Strategy Analyst specializing in matching candidate profiles to job requirements and estimating compensation conservatively.

Your task is to analyze the provided [Job Description] and [User Profile] (containing work history, education, and achievements) and output a single valid JSON object with exactly three keys: "salary_low", "salary_high", and "fit_rating".

Use concise and clear reasoning while thinking, using no more than 1000 tokens. Do not second-guess yourself

### Rules for Calculation:

1. **Salary Range ("salary_low" and "salary_high"):**
   - **Source Constraint:** Use ONLY the compensation information explicitly stated in the [Job Description]. Do not use external market data, inflation rates, or general industry averages unless the JD mentions them directly.
   - **Logic:** 
     - If the JD lists a range (e.g., "$80k - $100k"), return those exact numbers (in integers).
     - If the JD states a base salary (e.g., "$90k"), treat this as the minimum. Set "salary_low" to this value. Set "salary_high" to this value plus 5% to account for potential negotiation, capped at the next logical bracket if implied.
   - **Conservatism:** Never inflate the numbers. If the JD says "up to $120k", do not assume the average is $120k; estimate the range based on realistic distribution (e.g., low 95000, high 115000).
   - **Formatting:** Give estimated values as integers representing the annual compensation

2. **Fit Rating ("fit_rating"):**
   - **Scale:** Integer between 0 and 100.
   - **Methodology:** 
     - Analyze hard skills, years of experience, and education alignment against the JD's "Must-Haves."
     - Analyze soft skills and achievements against the JD's "Nice-to-Haves" and culture fit.
     - Penalize heavily for missing critical technical requirements.
     - Reward significantly for over-qualification and proven relevant achievements.
     - A score of 0 means zero match; 100 means a perfect match with all criteria exceeded.

### Output Format Constraints:
- Return ONLY the raw JSON string.
- Do NOT include markdown code blocks (```json ... ```).
- Do NOT include any introductory text, explanations, or conversational filler.
- Ensure the JSON is syntactically valid (proper quotes, commas, and braces).

### Example Output:
{
    "salary_low": 60000,
    "salary_high": 80000,
    "fit_rating": 87
}
"""

CREATE_RESUME_SYSTEM_PROMPT = """
You are an expert resume builder who specializes in taking a large amount of data and distilling it to a handful of extremely effective and efficient lines.

Your task is to analyze the Job Description and the Candidate Summary and output a single, valid JSON with the following keys: "name", "contact_info", "summary", "work_history", "education", "skills". Additional, optional keys can be included. In that case, the key will correspond to a resume section with a section name equal to the key
{
"name": str,
"contact_info": list[str],
"summary": str,
"work_history": list[dict],
"education": list[dict],
"skills": list[str],
otherkeys: list[str]
}

## name
a string with the Candidate's full name

## contact info
a list containing strings, where each string corresponds to exactly one piece of the candidate's contact info. These might be phone number, email address, location (give as City, ST), github. Give each of these pieces of information in a readable format

## summary
A single string comprised of no more than 2 sentences giving the most powerful personal summary possible for the job summary while still maintaining accuracy with the candidate's profile. If possible, include a relevant job title.

## work history
a list of dicts where each dict contains information relevant to a single company in the candidate's work history. These should be listed in reverse chronological order
This dict has the format of
{
"institution": str,
"dates": str,
"subtitle": str,
"location": str,
"bullets": list[str]
}
### institution
company name
### dates
string giving date range as abbreviated Month Year - month year (e.g. Apr. 2018 - Jun 2023)
### subtitle
string contaning Job title
### location
string containing city and state. If no location is given or it is a remote position, leave the location empty
### bullets
list containing relevant bullets, if any. The most relvant job can have up to 4 bullets. Each other job can have a maximum of 2

## education
a list of dicts where each dict contains information relevant to a single degree in the candidate's education history. These should be listed in reverse chronological order
This dict has the format of
{
"institution": str,
"dates": str,
"subtitle": str,
"location": str,
"bullets": list[str]
}
### institution
school name
### dates
string giving date range as abbreviated Month Year - month year (e.g. Apr. 2018 - Jun 2023)
### subtitle
string contaning degree info and (if applicable) GPA (e.g. BS Engineering, GPA: 3.5/4.0)
### location
string containing city and state. If no location is given or it is a remote position, leave the location empty
### bullets
list containing relevant bullets, if any. Scholarships, highly relevant club participation of classes can be put here. If work history is more relevant, then only keep one or two bullets total for the entire education history 

## skills
a list containing all of the relevant skills given by category and skill name. Category should be **bold**, and the individual skills within the category are comma separated plaintext (e.g. **Data entry**: Databases, Microsoft Excel, keyboarding, computer systems)

## other sections 
You can add one or two other sections that are relevant to the posting. These entries will contain a list of strings, where each string will correspond to a bullet under the section labeled by the key. Relevant sections might inlcude "Languages", "Publications", "Personal Projects", "Interests", "Projects & Presentations", "Projects & Publications", etc. Any markdown formatting applied to these bullets will correspondingly be applied in the final document

## Example Output
{
"name": "John Smith",
"contact_info": ["+1 (212) 123-4567", "jonny.smith@gmail.com", "https://linkedin.com/in/smith_john", "NYC, NY"],
"summary": "Senior Quality Engineer with 6+ years of experience leading teams at ABC Development. Skilled at autotesting workflows and quality control management.",
"work_history": [{
        "institution": "ABC Development",
        "dates": "Feb. 2015 - Jun. 2021",
        "subtitle": "Senior Quality Engineer",
        "location": "NYC, NY",
        "bullets": ["Enforced strict quality measures","Led a team of 17 individuals", "Routinely caught at least 6 critical errors each month"]
    },
    {
        "institution": "DEF Engineering",
        "dates": "Oct. 2009 - Jan. 2015",
        "subtitle": "Quality Enginer",
        "location": "Jersey City, NJ",
        "bullets": ["Developed auto-testing protocols", "Assisted in rigorous quality testing daily"]
    }],
"education": [{
        "institution": "FGH University",
        "dates": "Sep. 2005 - Jun. 2009",
        "subtitle": "BS in Mechanical Engineering, GPA: 3.6/4.0",
        "location": "Jersey City, NJ",
        "bullets": []
    }],
"skills":
    [
        "**Quality engineering**: Automated testing, mechanical safety",
        "**Team management**: Trust building, team reviewing, communication",
        "**Professional**: MS Office, technical documentation, communication"
    ],
"certifications":
    [
        "Certified Quality Engineer (ASQ), Six Sigma Black Belt"
    ],
"interests":
    [
        "Gaming, weightlifting, skateboarding, baking"
    ]
}

## Final instructions
After the name, contact info, and summary, the resume sections can be put in any order. Put the sections in the optimal order to ensure the most impressive and relevant sections are put higher up in order to catch recruiters' eyes. Also be ruthless with which information you choose to cut. Space on a resume is at a premium and every line has to tell the recruiter that this candidate is perfect for the job. Only include extra sections if they meaningfully add to the resume. Avoid redundancy so that multiple strengths can be highlighted.
Importantly, remember this: 
"""
# You are in debug right now to see how the system handles very long jsons, so insert enough content to create a resume 5+ pages long with at least 10 additional sections, with made up information if necessary
app = FastAPI(title="Job Fit Analyzer")


class AnalysisRequest(BaseModel):
    job_description: str
    candidate_profile: str


class AnalysisResponse(BaseModel):
    salary_low: int
    salary_high: int
    fit_score: int


class CreateRequest(BaseModel):
    job_description: str
    candidate_profile: str
    convo_context: list
    special_instruction: str

class CreateResponse(BaseModel):
    model_json: dict

class RenderRequest(BaseModel):
    resume_info: dict
    settings: dict

@app.post("/api/analyze", response_model=AnalysisResponse)
async def analyze(request: AnalysisRequest):
    # Placeholder until we implement the actual analysis.

    messages = [
        {"role": "system", "content": JOBINFO_SYSTEM_PROMPT},
        {"role": "system", "content": "Candidate summary: " + request.candidate_profile},
        {"role": "system", "content": "Job description: " + request.job_description}
    ]
    payload = {
                "model": modelname,
                "messages": messages,
                "max_tokens": 6000,
                "temperature": 0.8
            }

    response = requests.post(f"{BASE_URL}/v1/chat/completions", json=payload)
    reply = response.json().get("choices", [{}])[0].get("message",{}).get("content", "No response")
    analysis = parse_ressponse_for_json(reply)
    if analysis:
        return AnalysisResponse(
            salary_low=analysis['salary_low'],
            salary_high=analysis['salary_high'],
            fit_score=analysis['fit_rating'],
        )

@app.post("/api/create-resume", response_model=CreateResponse)
async def create(request: CreateRequest):
    messages = [
        {
            "role": "system",
            "content": CREATE_RESUME_SYSTEM_PROMPT + request.special_instruction
        },
        {
            "role": "system",
            "content": "Candidate summary: " + request.candidate_profile
        },
        {
            "role": "system",
            "content": "Job description: " + request.job_description
        }
    ]

    for turn in request.convo_context:
        messages.append(turn)

    payload = {
        "model": modelname,
        "messages": messages,
        "max_tokens": 12000,
        "temperature": 0.8
    }

    response = requests.post(
        f"{BASE_URL}/v1/chat/completions",
        json=payload
    )

    reply = (
        response.json()
        .get("choices", [{}])[0]
        .get("message", {})
        .get("content", "No response")
    )

    resume_info = parse_ressponse_for_json(reply)

    if resume_info is None:
        raise HTTPException(
            status_code=500,
            detail="Unable to parse resume from LLM response."
        )
    messages.append({"role": "model", "content": reply})
    with open(f"messages_for_diagnostics.jsonl", 'w', encoding='utf-8') as outfile:
        for message in messages:
            json.dump(message, outfile)
            outfile.write('\n')

    return CreateResponse(model_json=resume_info)

@app.post("/api/render-resume")
async def render(request: RenderRequest):
    resume_info = request.resume_info
    if resume_info:
        resume = Document()

        create_settings = request.settings

        body_font = create_settings["font"]
        x_margin = create_settings["horizontal_margin"]
        y_margin = create_settings["vertical_margin"]
        line_spacing = create_settings["line_spacing"]
        styles: dict[str, dict[str,str|int]] = {
            'Resume Normal': {
                "font": body_font,
                "size": 12,
                "bold": None,
                "ital": None,
                "underline": None
            },
            'Resume Title': {
                "font": body_font,
                "size": 24,
                "bold": True,
                "ital": False,
                "underline": False
            },
            'Resume Subtitle': {
                "font": body_font,
                "size": 13,
                "bold": False,
                "ital": False,
                "underline": None
            },
            'Resume Heading 1': {
                "font": body_font,
                "size": 14,
                "bold": True,
                "ital": False,
                "underline": False
            },
            'Resume Heading 2': {
                "font": body_font,
                "size": 12,
                "bold": True,
                "ital": False,
                "underline": False
            },
            'Resume Heading 3': {
                "font": body_font,
                "size": 12,
                "bold": False,
                "ital": True,
                "underline": False
            },
            'Resume Section Break': {
                "font": body_font,
                "size": 4,
                "bold": None,
                "ital": None,
                "underline": None
            }
        }

        resume_interpreter.define_styles(resume, styles)
        resume_interpreter.set_margins(resume, y_margin, x_margin)
        resume_interpreter.generate_header(resume, resume_info['name'], resume_info['contact_info'])
        resume_interpreter.generate_text(resume, resume_info['summary'])
        for key in resume_info.keys():
            if key.lower() == "name" or key.lower() == "contact_info" or key.lower() == "summary":
                continue
            if isinstance(resume_info[key][0], dict):
                resume_interpreter.generate_experience_section(resume, key.replace("_", " "), resume_info[key])
            else:
                resume_interpreter.generate_simple_section(resume, key.replace("_", " "), resume_info[key]) 
        resume_interpreter.set_spacing(resume, line_spacing)

        buffer = BytesIO()
        resume.save(buffer)
        buffer.seek(0)

        return StreamingResponse(
            buffer,
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            headers={
                "Content-Disposition":
                    'attachment; filename="generated_resume.docx"'
            }
        )




app.mount(
    "/static",
    StaticFiles(directory=FRONTEND_DIR),
    name="static",
)


@app.get("/")
async def serve_frontend():
    return FileResponse(FRONTEND_DIR / "index.html")

def parse_ressponse_for_json(response: str) -> dict|None:
    json_match = re.search(r'\{.*\}', response, re.DOTALL)
    parsed_data = None
    if json_match:
        json_string = json_match.group(0)
        try:
            parsed_data = json.loads(json_string)
            return parsed_data
        except:
            return None
