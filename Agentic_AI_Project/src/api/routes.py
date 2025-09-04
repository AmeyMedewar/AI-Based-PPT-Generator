import os
import shutil
from fastapi import APIRouter, UploadFile, File, HTTPException,Query
from fastapi.responses import FileResponse

from src.agents.content_extractor import ContentExtractor
from src.agents.Silver.PPT_Content_Generator_Silver import PPTCotentGeneratorSilver
from src.agents.Silver.slide_generator_silver import SlideGeneratorSilver
from src.agents.Gold.PPT_Content_Generator_Gold import PPTCotentGeneratorGold
from src.agents.Gold.slide_generator_gold import SlideGeneratorGold
from src.agents.Platinum.PPT_Content_Generator_Platinum import PPTCotentGeneratorPlatinum
from src.agents.Platinum.slide_generator_platinum import SlideGeneratorPlatinum

router = APIRouter()

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "Output"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@router.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    """Upload a file to the server"""
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return {"filename": file.filename, "path": file_path}



@router.post("/generate_silver/")
async def generate_slides_silver(
    topic: str = Query(..., description="Presentation topic"),
    filename: str = Query(None, description="Optional uploaded file"),
    audience: str = Query("", description="Target audience"),
    purpose: str = Query("", description="Purpose of presentation"),
    num_slides: int = Query(5, description="Number of slides"),
    style: str = Query("", description="Style/Tone"),
    complexity: str = Query("", description="Content complexity"),
    language: str = Query("", description="Language"),
    notes: str = Query("", description="Additional notes")
):
    """Generate slides with optional file input and content trimming"""
    
    # ---------------------------
    # 1. Extract text if file is provided
    # ---------------------------
    text = ""
    warning_message = ""

    if filename:
        file_path = os.path.join(UPLOAD_DIR, filename)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")

        extractor = ContentExtractor(directory=UPLOAD_DIR)
        text = extractor.extract_file(file_path)
        MAX_TOKENS = 500000
        # Trim text if it exceeds MAX_TOKENS
        words = text.split()
        if len(words) > MAX_TOKENS:
            text = " ".join(words[:MAX_TOKENS])
            warning_message = (
                f"⚠ Content exceeded {MAX_TOKENS} words. "
                f"Only the first {MAX_TOKENS} words were used for PPT generation."
            )

    # ---------------------------
    # 2. Generate content using PPTContentGenerator
    # ---------------------------
    content_generator = PPTCotentGeneratorSilver()
    summary = content_generator.generate_content(
        topic=topic,
        text=text,
        num_slides=num_slides,
        style=style or "Professional, concise",
        complexity=complexity or "Intermediate",
        audience=audience or "General Audience",
        purpose=purpose or "Informative Presentation",
        language=language or "English",
        notes=notes or ""
    )

    # ---------------------------
    # 3. Generate slides
    # ---------------------------
    slidegen = SlideGeneratorSilver(output=OUTPUT_DIR)
    slides_data = slidegen.parse_model_output(summary)

    # Fallback if parsing fails
    if not slides_data:
        bullet_points = summary.split("\n")  # simple fallback split
        slides_data = [{"title": topic, "bullets": bullet_points}]

    # Create the PowerPoint deck
    pptx_path = slidegen.create_slide_deck(slides_data=slides_data)

    # ---------------------------
    # 4. Prepare response
    # ---------------------------
    response = {
        "message": "Slides generated successfully",
        "pptx_path": pptx_path
    }

    if warning_message:
        response["warning"] = warning_message

    return response


@router.post("/generate_gold/")
async def generate_slides_gold(
    topic: str = Query(..., description="Presentation topic"),
    filename: str = Query(None, description="Optional uploaded file"),
    audience: str = Query("", description="Target audience"),
    purpose: str = Query("", description="Purpose of presentation"),
    num_slides: int = Query(5, description="Number of slides"),
    style: str = Query("", description="Style/Tone"),
    complexity: str = Query("", description="Content complexity"),
    language: str = Query("", description="Language"),
    notes: str = Query("", description="Additional notes")
):
    """Generate slides with optional file input and content trimming"""
    
    # ---------------------------
    # 1. Extract text if file is provided
    # ---------------------------
    text = ""
    warning_message = ""

    if filename:
        file_path = os.path.join(UPLOAD_DIR, filename)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")

        extractor = ContentExtractor(directory=UPLOAD_DIR)
        text = extractor.extract_file(file_path)
        MAX_TOKENS = 750000

        # Trim text if it exceeds MAX_TOKENS
        words = text.split()
        if len(words) > MAX_TOKENS:
            text = " ".join(words[:MAX_TOKENS])
            warning_message = (
                f"⚠ Content exceeded {MAX_TOKENS} words. "
                f"Only the first {MAX_TOKENS} words were used for PPT generation."
            )

    # ---------------------------
    # 2. Generate content using PPTContentGenerator
    # ---------------------------
    content_generator = PPTCotentGeneratorGold()
    summary = content_generator.generate_content(
        topic=topic,
        text=text,
        num_slides=num_slides,
        style=style or "Professional, concise",
        complexity=complexity or "Intermediate",
        audience=audience or "General Audience",
        purpose=purpose or "Informative Presentation",
        language=language or "English",
        notes=notes or ""
    )

    # ---------------------------
    # 3. Generate slides
    # ---------------------------
    slidegen = SlideGeneratorGold(output=OUTPUT_DIR)
    slides_data = slidegen.parse_model_output(summary)

    # Fallback if parsing fails
    if not slides_data:
        bullet_points = summary.split("\n")  # simple fallback split
        slides_data = [{"title": topic, "bullets": bullet_points}]

    # Create the PowerPoint deck
    pptx_path = slidegen.create_slide_deck(slides_data=slides_data)

    # ---------------------------
    # 4. Prepare response
    # ---------------------------
    response = {
        "message": "Slides generated successfully",
        "pptx_path": pptx_path
    }

    if warning_message:
        response["warning"] = warning_message

    return response



@router.post("/generate_platinum/")
async def generate_slides_platinum(
    topic: str = Query(..., description="Presentation topic"),
    filename: str = Query(None, description="Optional uploaded file"),
    audience: str = Query("", description="Target audience"),
    purpose: str = Query("", description="Purpose of presentation"),
    num_slides: int = Query(5, description="Number of slides"),
    style: str = Query("", description="Style/Tone"),
    complexity: str = Query("", description="Content complexity"),
    language: str = Query("", description="Language"),
    notes: str = Query("", description="Additional notes")
):
    """Generate slides with optional file input and content trimming"""
    
    # ---------------------------
    # 1. Extract text if file is provided
    # ---------------------------
    text = ""
    warning_message = ""

    if filename:
        file_path = os.path.join(UPLOAD_DIR, filename)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")

        extractor = ContentExtractor(directory=UPLOAD_DIR)
        text = extractor.extract_file(file_path)
        MAX_TOKENS = 750000

        # Trim text if it exceeds MAX_TOKENS
        words = text.split()
        if len(words) > MAX_TOKENS:
            text = " ".join(words[:MAX_TOKENS])
            warning_message = (
                f"⚠ Content exceeded {MAX_TOKENS} words. "
                f"Only the first {MAX_TOKENS} words were used for PPT generation."
            )

    # ---------------------------
    # 2. Generate content using PPTContentGenerator
    # ---------------------------
    content_generator = PPTCotentGeneratorPlatinum()
    summary = content_generator.generate_content(
        topic=topic,
        text=text,
        num_slides=num_slides,
        style=style or "Professional, concise",
        complexity=complexity or "Intermediate",
        audience=audience or "General Audience",
        purpose=purpose or "Informative Presentation",
        language=language or "English",
        notes=notes or ""
    )

    # ---------------------------
    # 3. Generate slides
    # ---------------------------
    slidegen = SlideGeneratorPlatinum(output=OUTPUT_DIR)
    slides_data = slidegen.parse_model_output(summary)

    # Fallback if parsing fails
    if not slides_data:
        bullet_points = summary.split("\n")  # simple fallback split
        slides_data = [{"title": topic, "bullets": bullet_points}]

    # Create the PowerPoint deck
    pptx_path = slidegen.create_slide_deck(slides_data=slides_data)

    # ---------------------------
    # 4. Prepare response
    # ---------------------------
    response = {
        "message": "Slides generated successfully",
        "pptx_path": pptx_path
    }

    if warning_message:
        response["warning"] = warning_message

    return response

@router.get("/download/")
async def download_pptx():
    """Download the generated PowerPoint file"""
    pptx_path = os.path.join(OUTPUT_DIR, "generated_ppt.pptx")
    if not os.path.exists(pptx_path):
        raise HTTPException(status_code=404, detail="No PPTX generated yet")
    return FileResponse(pptx_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="generated.pptx")
