import os
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.prompts import PromptTemplate

# Load API key from .env
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

if not api_key:
    raise ValueError("❌ GOOGLE_API_KEY not found. Please set it in your .env file")


class PPTCotentGeneratorSilver:
    def __init__(self):
        # Initialize Gemini model
        self.llm = ChatGoogleGenerativeAI(
            model="gemini-1.5-flash",
            google_api_key=api_key,
            temperature=0.3  # lower = more focused summaries
        )

        # Define a reusable summarization prompt
        self.prompt = PromptTemplate(
        input_variables=[
        "topic",
        "text",
        "num_slides",
        "style",
        "complexity",
        "audience",
        "purpose",
        "language",
        "notes"
            ],
        template="""
You are an expert presentation designer. 
Your task is to generate structured slide content for a PowerPoint presentation.

📌 Requirements:
- Topic: {topic}
- Target Audience: {audience}
- Purpose: {purpose}
- Number of Slides: {num_slides}
- Content Complexity: {complexity}
- Style/Tone: {style}
- Language: {language}
- Additional Notes: {notes}

📌 Input Material:
{text}

📌 Output Instructions:
**Refer the Input Material and give more importance to it if the topic is irrelevent from the input material leave the topic and use input material only and if the input material is not provided then use topic but overall give most importance to input material for output generation 
1. The **first slide must always be a Title Slide**, with:
   - Presentation title (use {topic})
   - Subtitle if relevant (purpose or audience)
   - [Optional: suggest a tagline or engaging subtitle]
2. From the second slide onwards, create slide content in clear **bullet points**.
3. Keep each slide focused (3–5 bullets max).
4. Suggest slide titles where appropriate.
5. Maintain the requested style, tone, and complexity.
6. Do not repeat the input verbatim — summarize and adapt for slides.
7. Return the result as a structured list like:

Slide 1: Title Slide
- Title: {topic}
- Subtitle: {purpose}
- [Optional: Tagline]

Slide 2: [Title]
- Bullet 1
- Bullet 2
...
    """,
)


    def generate_content(
        self,
        text: str,
        topic: str,
        num_slides: int,
        style: str,
        complexity: str,
        audience: str,
        purpose: str,
        language: str,
        notes: str
    ) -> str:
        chain = self.prompt | self.llm
        response = chain.invoke({
            "text": text,
            "topic": topic,
            "num_slides": num_slides,
            "style": style,
            "complexity": complexity,
            "audience": audience,
            "purpose": purpose,
            "language": language,
            "notes": notes,
        })
        return response.content


if __name__ == "__main__":
    # Example usage
    text = """
    Artificial Intelligence (AI) is transforming industries by automating tasks,
    improving decision-making, and enabling new products and services.
    Key applications include natural language processing, computer vision,
    and predictive analytics. While AI brings opportunities, it also raises
    concerns such as job displacement, bias in decision-making, and data privacy.
    """

    content_generator = PPTCotentGeneratorSilver()
    summary = content_generator.generate_content(
        text=text,
        topic="AI in Industry",
        num_slides=5,
        style="Professional, concise",
        complexity="Intermediate",
        audience="Business Executives",
        purpose="Awareness Presentation",
        language="English",
        notes="End with key takeaways and call-to-action"
    )
    print("\n✅ Generated Slide Content:\n", summary)
