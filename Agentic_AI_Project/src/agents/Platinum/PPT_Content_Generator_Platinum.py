import os
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.prompts import PromptTemplate

# Load API key from .env
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

if not api_key:
    raise ValueError("❌ GOOGLE_API_KEY not found. Please set it in your .env file")


class PPTCotentGeneratorPlatinum:
    def __init__(self):
        # Initialize Gemini model
        self.llm = ChatGoogleGenerativeAI(
            model="gemini-1.5-flash",
            google_api_key=api_key,
            temperature=0.3  # lower = more focused summaries
        )

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
Generate structured and intelligent slide content for a PowerPoint presentation based on the details provided.

📌 Presentation Context:
- **Topic**: {topic}
- **Target Audience**: {audience}
- **Purpose**: {purpose}
- **Number of Slides**: {num_slides}
- **Content Complexity**: {complexity}
- **Style/Tone**: {style}
- **Language**: {language}
- **Additional Notes**: {notes}

📌 Input Material:
{text}

📌 Output Instructions:

🎯 **Content Prioritization**:
- Always **prioritize the Input Material**. If the topic is irrelevant to the input, **ignore the topic** and use the input material.
- If no input material is provided, then use the topic alone.

🧩 **Slide Structuring Rules**:

1. **Slide 1: Title Slide**
   - **Title**: {topic}
   - **Subtitle**: {purpose} or {audience}
   - **Author**: Amey G M
   - **Date**: <current date>
   - _[Optional]_ **Tagline**: Engaging subtitle or phrase aligned with purpose

2. **Content Slides (Slide 2 onward)**
   - Use meaningful **slide titles**.
   - Each slide should contain **3–5 bullet points**, medium to large in length.
   - Bullets may include:
     - **Bold** for key terms
   - Subpoints may be used when appropriate.
   - Structure the slides logically to flow from introduction → explanation → elaboration.

3. **Optional: Question Slide (Insert if applicable)**
   - If questions are naturally suggested by the content, add a slide titled:
     - _“Questions to Consider”_ or _“Discussion Points”_
   - Format: 3–5 reflective or thought-provoking questions.

4. **Mandatory: Summary Slide (Always Second-to-Last)**
   - Title: **Summary / Key Takeaways**
   - 4–6 concise bullets summarizing the entire presentation.
   - Emphasize the most important messages, actions, or conclusions.
   - Use formatting to highlight priorities.

5. **Special Slide (If Applicable, Comes After Summary)**
   - Auto-detect and insert one of the following if relevant:
     - ⏳ **Timeline Slide**:
       - A 5-point chronological sequence
       - Format:
         - Year/Step 1: Description
         - ...
     - 🔄 **Flow Slide**:
       - A 5-point process or step-by-step flow
       - Format:
         - Step 1 → Step 2 → Step 3 → Step 4 → Step 5
     - ⚖️ **Comparison Slide**:
       - Highlight differences between two items or concepts
       - Format:
         - **Aspect 1**: A - ..., B - ...
         - ...

6. **Final Slide: Thank You Slide**
   - Message: _“Thank you for your attention”_
   - Optional closing statement or contact detail (if purpose supports it)

📌 Output Format:
Slide 1: Title Slide  
- Title: {topic}  
- Subtitle: {purpose}  
- Author: Amey G M  
- Date: <current date>  
- Tagline (if applicable): _[Tagline]_  

Slide 2: [Slide Title]  
- **Bullet 1**: Explanation  
- **Bullet 2**: Explanation  
...

[Optional] Slide X: Questions to Consider  
- _What is the impact of X on Y?_  
- _How does A differ from B in real-world application?_  
...

Slide N-2: Summary / Key Takeaways  
- **Key Point 1**  
- **Key Point 2**  
- **Key Point 3**  
...

[If applicable] Slide N-1: [Timeline / Flow / Comparison]  
- Format as described above

📌 Additional Notes:
- Avoid copying the input material verbatim — summarize and adapt.
- Maintain logical progression and reader-friendly formatting.
- Match the requested style, tone, and complexity throughout.
-Focus on generating special slides if possible
"""
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
    text=''

    content_generator = PPTCotentGeneratorPlatinum()
    summary = content_generator.generate_content(
        text=text,
        topic="ML vs DL",
        num_slides=5,
        style="Professional, concise",
        complexity="Intermediate",
        audience="Business Executives",
        purpose="Awareness Presentation",
        language="English",
        notes="End with key takeaways and call-to-action"
    )
    print("\n✅ Generated Slide Content:\n", summary)
