from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
import re
from collections import Counter
from langdetect import detect, LangDetectException
import langid
from lingua import Language, LanguageDetectorBuilder
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from googletrans import Translator as GoogleTranslator
from html.parser import HTMLParser
import shutil
from pathlib import Path

# Initialize FastAPI app
app = FastAPI(title="EZIcheck API")

# Configure CORS (allows your frontend to communicate with backend)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============= Your Translation Validation Logic =============

# Precompiled regex for numeric amounts.
numeric_pattern = re.compile(r"^[\$€£]?\s*\d+(?:[.,]\d+)?\s*$")

# Precompiled emoji regex pattern.
emoji_pattern = re.compile("[" 
    u"\U0001F600-\U0001F64F"  # emoticons
    u"\U0001F300-\U0001F5FF"  # symbols & pictographs
    u"\U0001F680-\U0001F6FF"  # transport & map symbols
    u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
    "]+", flags=re.UNICODE)

def extract_emojis(text):
    if not isinstance(text, str):
        text = str(text)
    return set(emoji_pattern.findall(text))

# Set of common self-closing HTML tags.
SELF_CLOSING_TAGS = {"area", "base", "br", "col", "command", "embed", "hr", "img", 
                     "input", "keygen", "link", "meta", "param", "source", "track", "wbr"}

# HTMLValidator using Python's HTMLParser.
class HTMLValidator(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tag_stack = []
        self.error = False
    def handle_starttag(self, tag, attrs):
        if tag not in SELF_CLOSING_TAGS:
            self.tag_stack.append(tag)
    def handle_endtag(self, tag):
        if tag in SELF_CLOSING_TAGS:
            return
        if not self.tag_stack or self.tag_stack[-1] != tag:
            self.error = True
        else:
            self.tag_stack.pop()
    def validate(self, text):
        self.tag_stack = []
        self.error = False
        try:
            self.feed(text)
            self.close()
        except Exception:
            return False
        return not self.error and not self.tag_stack

# Helper functions
def is_numeric(text):
    if not isinstance(text, str):
        text = str(text)
    return bool(numeric_pattern.match(text.strip()))

def is_all_hashtag_phrases(text):
    tokens = text.split()
    return bool(tokens) and all(token.startswith("#") and token.endswith("#") for token in tokens)

def clean_text(text):
    if not isinstance(text, str):
        text = str(text)
    EXCLUDE = {"Plus500", "+Premium"}
    return [w for w in text.split() if w not in EXCLUDE and not (w.startswith("#") and w.endswith("#"))]

def add_comment_to_cell(cell, comment_text):
    """Add or append a comment to a cell."""
    if cell.comment:
        existing_text = cell.comment.text
        cell.comment = Comment(f"{existing_text}\n{comment_text}", "Translation Validator")
    else:
        cell.comment = Comment(comment_text, "Translation Validator")
    cell.comment.width = 300
    cell.comment.height = 100

# Set up language detectors
supported_languages = list(Language.all())
lingua_detector = LanguageDetectorBuilder.from_languages(*supported_languages).build()
google_translator = GoogleTranslator()

def detect_language(text):
    results = []
    try:
        results.append(detect(text))
    except LangDetectException:
        results.append(None)
    try:
        results.append(langid.classify(text)[0])
    except Exception:
        results.append(None)
    try:
        lingua_result = lingua_detector.detect_language_of(text)
        results.append(lingua_result.iso_code_639_1.lower())
    except Exception:
        results.append(None)
    try:
        detection = google_translator.detect(text)
        results.append(detection.lang.lower())
    except Exception:
        results.append(None)
    return results

def detect_language_lingua_only(text):
    """Detect language using only Lingua detector."""
    try:
        lingua_result = lingua_detector.detect_language_of(text)
        return lingua_result.iso_code_639_1.lower() if lingua_result else None
    except Exception:
        return None

def process_excel_file(file_path: str, output_path: str) -> dict:
    """Process the Excel file and return statistics."""
    # Caches
    detection_cache = {}
    html_cache = {}
    lingua_only_cache = {}
    
    # Statistics
    stats = {
        "total_cells_checked": 0,
        "errors_found": 0,
        "warnings_found": 0,
        "sheets_processed": 0,
        "error_types": {
            "emoji": 0,
            "html": 0,
            "hashtag": 0,
            "language": 0,
            "untranslated": 0
        }
    }
    
    # Load workbook
    wb = load_workbook(file_path)
    
    # Define cell fills
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    
    # Process each worksheet
    for sheet in wb.worksheets:
        if sheet.title.strip().lower() == "parameters":
            continue
            
        stats["sheets_processed"] += 1
        max_col = sheet.max_column
        max_row = sheet.max_row
        
        for col_idx in range(4, max_col + 1):
            header_value = sheet.cell(row=1, column=col_idx).value
            if not header_value:
                continue
                
            lang_code = str(header_value).strip().lower()
            if lang_code == "cn":
                expected_lang = "zh_hans"
            elif lang_code == "zh":
                expected_lang = "zh_hant"
            elif lang_code == "nb_no":
                expected_lang = "no"
            else:
                expected_lang = lang_code
                
            for row_idx in range(2, max_row + 1):
                cell = sheet.cell(row=row_idx, column=col_idx)
                target_text = cell.value
                source_text = sheet.cell(row=row_idx, column=2).value
                target_str = str(target_text).strip() if target_text is not None else ""
                source_str = str(source_text).strip() if source_text is not None else ""
                
                stats["total_cells_checked"] += 1
                
                # Clear any existing comments
                cell.comment = None
                
                # Skip conditions
                if isinstance(target_text, str) and target_str.lower() == "#cfd_service#. #short_rw#":
                    continue
                
                if target_text is not None and source_text is not None:
                    if ((isinstance(target_text, (int, float)) or (isinstance(target_text, str) and is_numeric(target_text))) and
                        (isinstance(source_text, (int, float)) or (isinstance(source_text, str) and is_numeric(source_text)))):
                        if target_str.lower() == source_str.lower():
                            continue
                
                if target_str and is_all_hashtag_phrases(target_str) and target_str.lower() == source_str.lower():
                    continue
                
                # Emoji validation
                source_emojis = extract_emojis(source_str)
                if source_emojis:
                    target_emojis = extract_emojis(target_str)
                    if not source_emojis.issubset(target_emojis):
                        cell.fill = red_fill
                        missing_emojis = source_emojis - target_emojis
                        add_comment_to_cell(cell, f"EMOJI ERROR: Missing emojis from source: {', '.join(missing_emojis)}")
                        stats["errors_found"] += 1
                        stats["error_types"]["emoji"] += 1
                        continue
                
                # HTML validation
                if "<" in target_str and ">" in target_str:
                    if target_str in html_cache:
                        valid_html = html_cache[target_str]
                    else:
                        validator = HTMLValidator()
                        valid_html = validator.validate(target_str)
                        html_cache[target_str] = valid_html
                    if not valid_html:
                        cell.fill = red_fill
                        add_comment_to_cell(cell, "HTML ERROR: Invalid or unbalanced HTML tags detected")
                        stats["errors_found"] += 1
                        stats["error_types"]["html"] += 1
                        continue
                
                # Hashtag validation
                source_hashtags = re.findall(r"#(.*?)#", source_str)
                if source_hashtags:
                    target_hashtags = re.findall(r"#(.*?)#", target_str)
                    source_counter = Counter(source_hashtags)
                    target_counter = Counter(target_hashtags)
                    if source_counter != target_counter:
                        cell.fill = red_fill
                        stats["errors_found"] += 1
                        stats["error_types"]["hashtag"] += 1
                        
                        # Detailed error message
                        missing_from_target = []
                        extra_in_target = []
                        count_mismatches = []
                        
                        all_hashtags = set(source_hashtags) | set(target_hashtags)
                        for hashtag in all_hashtags:
                            source_count = source_counter.get(hashtag, 0)
                            target_count = target_counter.get(hashtag, 0)
                            if source_count > target_count:
                                if target_count == 0:
                                    missing_from_target.append(f"#{hashtag}#")
                                else:
                                    count_mismatches.append(f"#{hashtag}# (expected {source_count}, found {target_count})")
                            elif target_count > source_count:
                                if source_count == 0:
                                    extra_in_target.append(f"#{hashtag}#")
                                else:
                                    count_mismatches.append(f"#{hashtag}# (expected {source_count}, found {target_count})")
                        
                        error_messages = []
                        if missing_from_target:
                            error_messages.append(f"Missing: {', '.join(missing_from_target)}")
                        if extra_in_target:
                            error_messages.append(f"Extra: {', '.join(extra_in_target)}")
                        if count_mismatches:
                            error_messages.append(f"Count mismatch: {', '.join(count_mismatches)}")
                        
                        add_comment_to_cell(cell, f"HASHTAG ERROR: {'; '.join(error_messages)}")
                        continue
                
                # Untranslated check
                if not target_str or target_str.lower() == source_str.lower():
                    if source_str:
                        cell.fill = grey_fill
                        stats["warnings_found"] += 1
                        stats["error_types"]["untranslated"] += 1
                        if not target_str:
                            add_comment_to_cell(cell, "UNTRANSLATED: Empty translation")
                        else:
                            add_comment_to_cell(cell, "UNTRANSLATED: Target text matches source text")
                    continue
                
                # Language detection
                cleaned_words = clean_text(target_str)
                combined_text = " ".join(cleaned_words)
                
                if len(cleaned_words) < 4:
                    # Short text - use Lingua only
                    if combined_text in lingua_only_cache:
                        detected_language = lingua_only_cache[combined_text]
                    else:
                        detected_language = detect_language_lingua_only(combined_text)
                        lingua_only_cache[combined_text] = detected_language
                    
                    if detected_language is not None:
                        if expected_lang in ("zh_hans", "zh_hant"):
                            match = detected_language == "zh"
                        else:
                            match = detected_language == expected_lang
                        
                        if not match:
                            cell.fill = red_fill
                            stats["errors_found"] += 1
                            stats["error_types"]["language"] += 1
                            add_comment_to_cell(cell, 
                                f"LANGUAGE ERROR (short text): Expected '{expected_lang}', "
                                f"but Lingua detected '{detected_language}' "
                                f"(text has {len(cleaned_words)} word(s) after cleaning)")
                else:
                    # Long text - use all detectors
                    if combined_text in detection_cache:
                        detected_languages = detection_cache[combined_text]
                    else:
                        detected_languages = detect_language(combined_text)
                        detection_cache[combined_text] = detected_languages
                    
                    if expected_lang in ("zh_hans", "zh_hant"):
                        match = any(lang == "zh" for lang in detected_languages if lang is not None)
                    else:
                        match = any(lang == expected_lang for lang in detected_languages if lang is not None)
                    
                    if not match and all(lang != expected_lang for lang in detected_languages if lang is not None):
                        cell.fill = red_fill
                        stats["errors_found"] += 1
                        stats["error_types"]["language"] += 1
                        
                        detector_names = ["langdetect", "langid", "Lingua", "Google Translate"]
                        detection_results = []
                        for i, lang in enumerate(detected_languages):
                            if lang is not None:
                                detection_results.append(f"{detector_names[i]}: {lang}")
                        
                        add_comment_to_cell(cell, 
                            f"LANGUAGE ERROR: Expected '{expected_lang}', "
                            f"but detected: {', '.join(detection_results)}")
    
    # Save the processed file
    wb.save(output_path)
    return stats

# ============= API Endpoints =============

@app.get("/")
async def root():
    return {"message": "EZIcheck API is running!", "version": "1.0.0"}

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

@app.post("/validate")
async def validate_excel(file: UploadFile = File(...)):
    """
    Validate an Excel file for translation errors.
    Returns the processed file and statistics.
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files are supported")
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Save uploaded file
        input_path = os.path.join(temp_dir, file.filename)
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        # Process file
        output_filename = f"validated_{file.filename}"
        output_path = os.path.join(temp_dir, output_filename)
        
        stats = process_excel_file(input_path, output_path)
        
        # Return processed file
        return FileResponse(
            path=output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_filename,
            headers={
                "X-Stats": str(stats),  # Include stats in headers
                "Access-Control-Expose-Headers": "X-Stats"  # Allow frontend to read this header
            }
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Cleanup will happen after file is sent
        # You might want to implement a cleanup task
        pass

@app.post("/validate-with-stats")
async def validate_excel_with_stats(file: UploadFile = File(...)):
    """
    Validate an Excel file and return both the file download URL and statistics.
    """
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files are supported")
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Save and process file
        input_path = os.path.join(temp_dir, file.filename)
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        output_filename = f"validated_{file.filename}"
        output_path = os.path.join(temp_dir, output_filename)
        
        stats = process_excel_file(input_path, output_path)
        
        # For this endpoint, we'd need to store the file temporarily
        # and return a download link. This is a simplified version.
        return {
            "filename": output_filename,
            "stats": stats,
            "message": "File processed successfully"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))