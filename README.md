# AI-Based-PPT-Generator
An AI-powered presentation generator that creates professional PowerPoint slides from just a prompt. You can customize based on content complexity, target audience, and presentation style.


Three levels of slide generation:

Basic 🟢 → Quick, text-based, simple slides.

Premium 🟡 → Polished design, structured content, relevant images.

Super Premium 🔴 → Executive-grade presentations with refined visuals, audience targeting, and advanced formatting.

📌 Table of Contents
1. Features
2. User Guide
3. Modes
4. Tech Stack
5. Developer Guide
6. API Endpoints
7. Contribution
8. License
9. Future Roadmap

🚀 Features

✅ Generate PPTs from natural language prompts
✅ Adjust for audience type (students, executives, researchers, etc.)
✅ Control content complexity (basic, intermediate, advanced)
✅ Choose mode → Basic, Premium, or Super Premium
✅ Auto-generate slide titles, bullets, and images
✅ Export directly to .pptx
✅ Optional AI-generated diagrams & visuals

👤 User Guide

1. Installation

git clone https://github.com/yourusername/ai-ppt-generator.git
cd ai-ppt-generator
pip install -r requirements.txt

Create a .env file with your keys:

OPENAI_API_KEY=your_key
HUGGINGFACE_API_KEY=your_key

2. Run
For Python backend:
python app.py
For Streamlit app:
streamlit run app.py

4. Output
Basic Mode → Text-focused slides.
Premium Mode → Polished formatting + AI images.
Super Premium Mode → Refined executive-style slides with optimized flow.


🛠 Tech Stack
Language: Python
Frameworks: Django / Flask / Streamlit (choose your actual one)
Frontend: React + TailwindCSS (if applicable)
PPT Generation: python-pptx
AI Models: OpenAI GPT / HuggingFace / Gemini
Image Generation: Stable Diffusion / DALL·E
Database: MongoDB / PostgreSQL (optional)
Deployment: Vercel / Docker / Heroku

👨‍💻 Developer Guide
Project Structure
ai-ppt-generator/
│── src/
│   ├── agents/           # AI logic & prompt processing
│   ├── ppt_builder/      # PPT generation scripts
│   ├── api/              # API endpoints
│── app.py                # Main entry point
│── requirements.txt
│── README.md

Environment Setup

python -m venv venv
source venv/bin/activate
pip install -r requirements.txt

🌐 API Endpoints

Method	Endpoint	Description
POST	/generate_ppt	Generate a PPT from prompt + parameters
GET	/download/{id}	Download a generated PPT
GET	/health	Health check endpoint

Example Request
POST /generate_ppt
{
  "prompt": "AI in Healthcare",
  "mode": "super_premium",
  "audience": "business_executives",
  "complexity": "medium",
  "slides": 12
}

Example Response

{
  "status": "success",
  "file_url": "/download/12345"
}

📝 License
This project is licensed under the MIT License.

🚀 Future Roadmap
🎤 Voice-to-Prompt → Generate slides from speech input
📊 Charts & Graphs → Auto-generate graphs from data
🌎 Multi-language support
🎨 Custom Themes → Brand-specific designs

