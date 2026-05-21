import json
import uuid
import threading
import os
from datetime import datetime
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
import google.genai as genai
from google.genai import types
from upstash_redis import Redis
from dotenv import load_dotenv

# Import the decoupled formatting engine from your local directory
from document_builder import build_docx_binary

# Load variables from your project's .env file automatically
load_dotenv()

# Initialize Cloud Connections from Environment
redis = Redis(url=os.environ.get("UPSTASH_REDIS_REST_URL"), token=os.environ.get("UPSTASH_REDIS_REST_TOKEN"))
client = genai.Client()

# Hardcoded Baseline Resume Data Configuration
VIGNESHWAR_BASE_RESUME = {
    "name": "VIGNESHWAR REDDY DONAPATI",
    "phone": "+91 9908731210",
    "email": "donapativigneshwar@gmail.com",
    "portfolio": "imvignesh.in",
    "linkedin": "linkedin.com/in/vignesh457",
    "github": "github.com/vignesh457",
    "skills": {
        "Programming Languages": ["C", "Java", "Python", "JavaScript", "TypeScript"],
        "Front End Technologies": ["HTML", "CSS", "React.js", "Redux", "Redux Toolkit", "Next.js", "React Native", "jQuery", "Tailwind CSS"],
        "Back End Technologies": ["Node.js", "Express.js", "EJS", "Spring MVC", "Spring Boot", "Docker", "Redis"],
        "Databases": ["MySQL", "Postgres", "Prisma", "Firebase", "MongoDB (NoSQL)"],
        "Additional Tools": ["Git", "GitHub", "Postman", "Jira", "JWT", "Visual Studio Code", "IntelliJ IDEA"],
        "Computer Fundamentals": ["Object oriented Programming", "DBMS", "Data structures & algorithms", "System Design (LLD, HLD)"]
    },
    "experience": [
        {
            "title": "Associate Software Engineering",
            "company": "Accenture",
            "dates": "Nov 2024 – Present",
            "bullets": [
                "Developed and maintained a scalable microservices application using Spring Boot, Node.js, and React to support critical import/export workflows and reliable cross-system data processing.",
                "Troubleshot and resolved complex production incidents by tracing requests across backend services, SQL databases, Azure Service Bus DLQ, and logs, reducing issue recurrence and improving overall system reliability.",
                "Collaborated with cross-functional teams during release cycles to drive root-cause analysis and code improvements, ensuring smooth Blue-Green deployments with minimal to zero downtime."
            ]
        },
        {
            "title": "Associate Software Engineering",
            "company": "Tech Mahindra",
            "dates": "Mar 2024 – Nov 2024",
            "bullets": [
                "Developed a full-stack internal productivity tool using React.js and Node.js to automate and streamline team workflows, resulting in improved operational efficiency and reduced manual effort.",
                "Implemented Role-Based Authentication (RBAC) and integrated secure internal APIs in Node.js, ensuring controlled data access and reliable data management across multiple user roles."
            ]
        }
    ],
    "projects": [
        {
            "title": "NumberHunt",
            "link": "https://github.com/vignesh457/number-hunt",
            "video_demo": "",
            "bullets": [
                "Built a multiplayer reflex-based game using React Native (Expo, TypeScript) with real-time gameplay via Socket.IO WebSocket, backend (Node.js + Express), supporting solo(offline) and multi-player(Room-based online) modes with animations, bgm, and seamless UI using Expo Router, NativeWind, and Moti.",
                "•	Designed and deployed a global leaderboard and user state system with Redux Toolkit, PostgreSQL, and Prisma ORM; packaged the app with Expo EAS and deployed the frontend on PhonePe Indus Appstore, ensuring optimized production-ready builds."
            ]
        },
        {
            "title": "ArtNook",
            "link": "https://artnook.vercel.app",
            "video_demo": "",
            "bullets": [
                "Developed a responsive full-stack Artist Showcase Platform designed to help independent artists display their portfolios, manage profiles, and reach a wider audience. Built using React.js, Node.js, and MongoDB, incorporating secure JWT authentication, protected routes, and role-based access control for artwork management.",
                "Improved application performance and usability through React memoization, lazy loading, advanced search, and category filtering features, delivering a smooth browsing experience for users."
            ]
        },
        {
            "title": "Spotify Clone",
            "link": "https://spotify-vignesh.netlify.app",
            "video_demo": "",
            "bullets": [
                "Built a responsive Spotify clone with HTML, CSS, and JavaScript with asynchronous programming.",
                "Implemented essential features like play/pause, previous/next controls, progress bar, volume adjustment, and search functionality using JavaScript’s DOM manipulation and event handling."
            ]
        }
    ],
    "certifications": [
        {
            "title": "Oracle Certified Associate, Java SE 8 Programmer",
            "link": "https://drive.google.com/file/d/1MeFkyClzrviKFX2MXWcz26PB7Obn1-Mk/view"
        }
    ],
    "education": [
        {
            "institution": "CMR INSTITUTE OF TECHNOLOGY, HYDERABAD",
            "degree": "Bachelor of Technology",
            "field": "Electronic and Communication Engineering",
            "dates": "August 2019 – July 2023",
            "gpa": "8.50"
        }
    ],
    "coding_profiles": [
        "LeetCode (1530, Top 25%, 350+ solved)",
        "GFG (580 score, 250+ solved)",
        "CodeChef (1554, 2 ★ , 126 contests)"
    ]
}

def background_tailor_worker(task_id, jd_txt):
    """Worker thread using Vigneshwar's baseline data to target a specific Job Description"""
    try:
        # Define a clean structural blueprint WITHOUT hardcoded placeholder values
        json_blueprint = """{
          "analytics": {
             "ats_score_before": "An integer between 0 and 100 representing the match score of the original baseline resume against the JD",
             "ats_score_after": "An integer between 0 and 100 representing the optimized match score after applying your keyword tailored adjustments",
             "modifications": [
                {"section": "The specific resume section modified", "description": "A brief explanation of what keywords or framing were optimized"}
             ]
          },
          "resume_data": {
              "name": "VIGNESHWAR REDDY DONAPATI",
              "phone": "+91 9908731210",
              "email": "donapativigneshwar@gmail.com",
              "portfolio": "imvignesh.in",
              "linkedin": "linkedin.com/in/vignesh457",
              "github": "github.com/vignesh457",
              "skills": {},
              "experience": [],
              "projects": [],
              "certifications": [],
              "education": [],
              "coding_profiles": []
          }
        }"""

        prompt = f"""
        You are an expert resume writer and software engineer ATS specialist.
        Task: Review the baseline resume JSON profile for Vigneshwar Reddy Donapati and tailor it to maximize match parameters against the provided Job Description (JD).
        
        CRITICAL SCORING MANDATE:
        1. Calculate a genuine, honest 'ats_score_before' by comparing how many core tech stack items, frameworks, and job requirements from the target JD are missing or named differently in the current baseline profile.
        2. Calculate a realistic 'ats_score_after' based on how closely your newly tailored version matches the exact target JD specifications. Do not simply default to 95; make it an accurate reflection of the alignment.
        
        CRITICAL ATS KEYWORD MATCHING RULES:
        1. Parse the target Job Description closely for literal spelling and punctuation preferences.
        2. Replicate the EXACT phrasing, spelling, capitalization, and punctuation variants used in the job post text.
           - If the JD writes "full-stack", do NOT write "Full Stack" or "full stack".
           - If the JD writes "ReactJS", do NOT write "React.js" or "React".
           - If the JD writes "backend", do NOT write "Back End" or "Back-end".
        3. Match these tokens exactly across your tailored skills and experience strings to ensure a perfect literal hit for strict ATS tokenizers.
        
        CRITICAL CONTENT DIRECTIONS:
        1. Keep ALL baseline technical items under the skills categories intact. Do not drop or abbreviate any frameworks.
        2. Extract and preserve real URLs/hyperlinks for projects, video demos, or code repositories.
        3. Extract and preserve the "Certifications" and "Coding Profiles" sections completely from the input data profile.
        
        Output Requirement: Return output exclusively in valid JSON matching this schema format structure:
        {json_blueprint}
        
        Baseline Profile Data: {json.dumps(VIGNESHWAR_BASE_RESUME)}
        Target Job Description: {jd_txt}
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json"),
        )
        
        payload = json.loads(response.text.strip())
        doc_hex = build_docx_binary(payload['resume_data'])
        
        # Calculate scores and append custom timestamp string values
        ats_score = payload['analytics'].get('ats_score_after', 90)
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        generated_filename = f"vigneshwar_{ats_score}_{timestamp_str}.docx"
        
        redis.set(f"analytics:{task_id}", json.dumps(payload['analytics']))
        redis.set(f"file:{task_id}", doc_hex)
        redis.set(f"filename:{task_id}", generated_filename)
        redis.set(f"status:{task_id}", "COMPLETED")
    except Exception as e:
        redis.set(f"status:{task_id}", f"FAILED: {str(e)}")

class handler(BaseHTTPRequestHandler):
    def do_HEAD(self):
        """Handle internal infrastructure health checks smoothly"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        content_length = int(self.headers['Content-Length'])
        post_data = json.loads(self.rfile.read(content_length).decode('utf-8'))
        
        task_id = str(uuid.uuid4())
        redis.set(f"status:{task_id}", "PROCESSING")
        
        # Notice: Only passing down jd input parameter string to background worker thread now
        worker = threading.Thread(
            target=background_tailor_worker, 
            args=(task_id, post_data.get('jd',''))
        )
        worker.start()
        
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"task_id": task_id}).encode('utf-8'))

    def do_GET(self):
        parsed = urlparse(self.path)
        query_components = parse_qs(parsed.query)
        task_id = query_components.get("task_id", [None])[0]
        download_trigger = query_components.get("download", [None])[0]
        path = parsed.path

        # If no task_id provided, treat as a health check or static file request
        if not task_id:
            # Serve the project's index.html for root requests when available
            if path in ('/', '/index.html'):
                try:
                    index_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'index.html'))
                    with open(index_path, 'r', encoding='utf-8') as fh:
                        content = fh.read()
                    self.send_response(200)
                    self.send_header('Content-Type', 'text/html; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(content.encode('utf-8'))
                    return
                except Exception:
                    # Fall through to return a simple JSON health response
                    pass

            # Generic health check response for probes without task_id
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "OK"}).encode('utf-8'))
            return

        status = redis.get(f"status:{task_id}")
        
        if status == "COMPLETED":
            if download_trigger == "true":
                file_hex = redis.get(f"file:{task_id}")
                filename = redis.get(f"filename:{task_id}") or f"vigneshwar_tailored_{task_id[:4]}.docx"
                file_bytes = bytes.fromhex(file_hex)
                
                self.send_response(200)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
                self.end_headers()
                self.wfile.write(file_bytes)
            else:
                analytics_data = json.loads(redis.get(f"analytics:{task_id}"))
                self.send_response(200)
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "COMPLETED", "analytics": analytics_data}).encode('utf-8'))
        else:
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": status or "NOT_FOUND"}).encode('utf-8'))

if __name__ == '__main__':
    from http.server import HTTPServer
    import os
    
    # Dynamically bind to the port Render gives your container instance
    port = int(os.environ.get("PORT", 8000))
    server = HTTPServer(('0.0.0.0', port), handler)
    print(f"Local test server running smoothly on port {port}")
    server.serve_forever()