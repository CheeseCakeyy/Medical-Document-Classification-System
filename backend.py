"""
backend.py — with batch processing support added
"""

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

import os
import re
import shutil
import logging
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import cv2
import pytesseract
from PIL import Image
import fitz

from PySide6.QtCore import QThread, Signal

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

CATEGORIES = [
    "prescription", "lab_report", "discharge_summary",
    "medical_bill", "insurance_claim", "referral_letter", "other",
]

CONFIDENCE_THRESHOLD = 0.70
DEFAULT_OUTPUT_DIR = Path("classified_documents")
SUPPORTED_EXTENSIONS = {".pdf", ".png", ".jpg", ".jpeg", ".tiff", ".tif", ".bmp", ".webp"}

PHRASE_BONUS = 3

RULES = {
    "prescription": [
        "prescription", "rx", "prescribed", "medication", "dosage", "dose",
        "mg", "tablet", "capsule", "syrup", "refill", "sig:",
        "twice daily", "once daily", "thrice daily", "after food", "before food",
        "dispense", "no refills", "doctor", "physician", "clinic", "pharmacy",
    ],
    "lab_report": [
        "lab report", "laboratory report", "pathology report", "test report",
        "blood test", "complete blood count", "cbc", "wbc", "rbc",
        "hemoglobin", "platelets", "hematocrit",
        "cholesterol", "triglycerides", "ldl", "hdl", "glucose", "hba1c",
        "creatinine", "urea", "bilirubin", "sgpt", "sgot", "tsh",
        "urinalysis", "urine culture", "stool test",
        "normal range", "reference range", "within normal limits",
        "scan report", "mri", "ct scan", "x-ray", "ultrasound", "ecg",
        "biopsy", "histopathology", "positive", "negative", "reagent",
    ],
    "discharge_summary": [
        "discharge summary", "discharge note",
        "date of admission", "date of discharge", "admission date", "discharge date",
        "hospital course", "discharged", "admitted on", "inpatient",
        "final diagnosis", "procedure performed",
        "discharge instructions", "condition at discharge",
        "follow up after", "follow-up in",
    ],
    "medical_bill": [
        "invoice", "bill", "receipt",
        "amount due", "total amount", "net payable", "balance due",
        "consultation fee", "procedure charges", "itemized bill",
        "insurance adjustment", "amount paid", "outstanding balance",
        "billing statement", "account statement",
        "subtotal", "tax", "gst", "discount",
    ],
    "insurance_claim": [
        "insurance claim", "claim form", "claim number",
        "policy number", "policy holder", "member id", "group number",
        "pre-authorization", "pre authorization", "authorization number",
        "explanation of benefits", "eob", "reimbursement", "cashless claim",
        "tpa", "third party administrator", "deductible", "copay",
        "icd code", "cpt code", "diagnosis code", "procedure code",
        "network provider", "in-network",
    ],
    "referral_letter": [
        "referral letter", "referral note",
        "dear dr", "dear doctor", "i am referring", "i am writing to refer",
        "kindly see this patient", "please see this patient",
        "for your evaluation", "for consultation", "for specialist opinion",
        "under my care", "my patient", "kindly advise",
        "further management", "requesting evaluation",
    ],
}


# ─────────────────────────────────────────────
# Rule-based Classifier
# ─────────────────────────────────────────────
class RuleClassifier:
    def load(self) -> Tuple[bool, str]:
        return True, "Rule-based classifier ready."

    @property
    def is_loaded(self) -> bool:
        return True

    def predict(self, text: str, lines: list, file_ext: str) -> Tuple[str, float, dict]:
        text_lower = text.lower()
        scores = {cat: 0 for cat in RULES}
        for category, keywords in RULES.items():
            for kw in keywords:
                if " " in kw:
                    scores[category] += text_lower.count(kw) * PHRASE_BONUS
                else:
                    scores[category] += len(re.findall(rf"\b{re.escape(kw)}\b", text_lower))
        total = sum(scores.values())
        if total == 0:
            n = len(RULES)
            uniform = round(1.0 / n, 4)
            return "other", uniform, {c: uniform for c in RULES}
        probs = {cat: round(s / total, 4) for cat, s in scores.items()}
        best_cat = max(probs, key=probs.get)
        return best_cat, probs[best_cat], probs


# ─────────────────────────────────────────────
# OCR Engine
# ─────────────────────────────────────────────
class OCREngine:
    @staticmethod
    def preprocess_image(image: np.ndarray) -> np.ndarray:
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        denoised = cv2.fastNlMeansDenoising(gray, h=10)
        thresh = cv2.adaptiveThreshold(
            denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
        )
        return thresh

    @staticmethod
    def extract_from_image(image_path: str) -> Tuple[str, list, bool]:
        try:
            img = cv2.imread(image_path)
            if img is None:
                pil_img = Image.open(image_path)
                img = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
            if img is None:
                return "", [], False
            preprocessed = OCREngine.preprocess_image(img)
            raw = pytesseract.image_to_string(preprocessed, config="--psm 6")
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            text = " ".join(lines)
            if not text:
                return "", [], False
            return text, lines, True
        except pytesseract.TesseractNotFoundError:
            logger.error("Tesseract not found.")
            return "", [], False
        except Exception as e:
            logger.error(f"Image OCR error: {e}")
            return "", [], False

    @staticmethod
    def extract_from_pdf(pdf_path: str) -> Tuple[str, list, bool]:
        try:
            doc = fitz.open(pdf_path)
            all_lines = []
            for page_num, page in enumerate(doc):
                native_text = page.get_text("text").strip()
                if native_text:
                    all_lines.extend([ln.strip() for ln in native_text.splitlines() if ln.strip()])
                else:
                    pix = page.get_pixmap(dpi=200)
                    img_data = np.frombuffer(pix.samples, dtype=np.uint8)
                    img = img_data.reshape(pix.height, pix.width, pix.n)
                    if pix.n == 4:
                        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
                    elif pix.n == 3:
                        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                    preprocessed = OCREngine.preprocess_image(img)
                    ocr_raw = pytesseract.image_to_string(preprocessed, config="--psm 6")
                    all_lines.extend([ln.strip() for ln in ocr_raw.splitlines() if ln.strip()])
            doc.close()
            text = " ".join(all_lines)
            if not text:
                return "", [], False
            return text, all_lines, True
        except Exception as e:
            logger.error(f"PDF extraction error: {e}")
            return "", [], False

    @classmethod
    def extract_text(cls, file_path: str) -> Tuple[str, list, bool]:
        ext = Path(file_path).suffix.lower()
        if ext == ".pdf":
            return cls.extract_from_pdf(file_path)
        elif ext in SUPPORTED_EXTENSIONS:
            return cls.extract_from_image(file_path)
        else:
            return "", [], False


# ─────────────────────────────────────────────
# File Organizer
# ─────────────────────────────────────────────
class FileOrganizer:
    def __init__(self, output_dir: Optional[str] = None):
        self.output_dir = Path(output_dir or DEFAULT_OUTPUT_DIR)
        self._ensure_folders()

    def _ensure_folders(self):
        for cat in CATEGORIES:
            (self.output_dir / cat).mkdir(parents=True, exist_ok=True)

    def save_document(self, source_path: str, category: str) -> Tuple[str, bool]:
        try:
            if category not in CATEGORIES:
                category = "other"
            dest_folder = self.output_dir / category
            dest_folder.mkdir(parents=True, exist_ok=True)
            src = Path(source_path)
            dest = dest_folder / src.name
            counter = 1
            while dest.exists():
                dest = dest_folder / f"{src.stem}_{counter}{src.suffix}"
                counter += 1
            shutil.copy2(source_path, dest)
            return str(dest), True
        except Exception as e:
            logger.error(f"File save error: {e}")
            return "", False

    def get_output_dir(self) -> str:
        return str(self.output_dir)


# ─────────────────────────────────────────────
# Document Classifier Facade
# ─────────────────────────────────────────────
class DocumentClassifier:
    def __init__(self, model_path: str = "model.pkl", output_dir: Optional[str] = None):
        self.model_manager = RuleClassifier()
        self.ocr_engine = OCREngine()
        self.file_organizer = FileOrganizer(output_dir)
        self._model_loaded = False

    def initialize(self) -> Tuple[bool, str]:
        success, msg = self.model_manager.load()
        self._model_loaded = success
        return success, msg

    def process_document(self, file_path: str) -> dict:
        result = {
            "success": False, "ocr_text": "", "lines": [],
            "category": "other", "confidence": 0.0, "all_probs": {},
            "needs_review": False, "saved_path": "", "error": "",
        }
        text, lines, ocr_ok = self.ocr_engine.extract_text(file_path)
        result["ocr_text"] = text
        result["lines"] = lines
        if not ocr_ok or not text.strip():
            result["error"] = "OCR failed or returned empty text."
            dest, _ = self.file_organizer.save_document(file_path, "other")
            result["saved_path"] = dest
            return result
        file_ext = Path(file_path).suffix.lower().lstrip(".")
        try:
            category, confidence, all_probs = self.model_manager.predict(text, lines, file_ext)
            result["category"] = category
            result["confidence"] = confidence
            result["all_probs"] = all_probs
            result["needs_review"] = confidence < CONFIDENCE_THRESHOLD
        except Exception as e:
            result["error"] = f"Classification error: {e}"
            result["category"] = "other"
        dest, saved = self.file_organizer.save_document(file_path, result["category"])
        result["saved_path"] = dest
        result["success"] = saved
        return result

    def get_output_directory(self) -> str:
        return self.file_organizer.get_output_dir()


# ─────────────────────────────────────────────
# Batch Worker (QThread)
# ─────────────────────────────────────────────
class BatchWorker(QThread):
    """
    Processes a list of files sequentially in a background thread.

    Signals:
        file_started(index, total, filename)   — emitted before each file
        file_done(index, total, result_dict)   — emitted after each file
        batch_finished(summary_dict)           — emitted when all files done
        batch_error(filename, error_msg)       — emitted on per-file error
    """
    file_started  = Signal(int, int, str)   # idx, total, filename
    file_done     = Signal(int, int, dict)  # idx, total, result
    batch_finished = Signal(dict)           # summary
    batch_error   = Signal(str, str)        # filename, error

    def __init__(self, classifier: DocumentClassifier, file_paths: list[str]):
        super().__init__()
        self.classifier = classifier
        self.file_paths = file_paths
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        total = len(self.file_paths)
        summary = {
            "total": total, "success": 0, "failed": 0,
            "needs_review": 0, "by_category": {c: 0 for c in CATEGORIES},
            "results": [],
        }

        for idx, file_path in enumerate(self.file_paths, start=1):
            if self._cancelled:
                break

            filename = Path(file_path).name
            self.file_started.emit(idx, total, filename)

            try:
                result = self.classifier.process_document(file_path)
                result["filename"] = filename
                result["file_path"] = file_path
                result["index"] = idx

                if result["success"]:
                    summary["success"] += 1
                    summary["by_category"][result["category"]] = (
                        summary["by_category"].get(result["category"], 0) + 1
                    )
                else:
                    summary["failed"] += 1

                if result.get("needs_review"):
                    summary["needs_review"] += 1

                summary["results"].append(result)
                self.file_done.emit(idx, total, result)

            except Exception as e:
                err_result = {
                    "filename": filename, "file_path": file_path,
                    "index": idx, "success": False, "error": str(e),
                    "category": "other", "confidence": 0.0,
                    "saved_path": "", "needs_review": False,
                }
                summary["failed"] += 1
                summary["results"].append(err_result)
                self.file_done.emit(idx, total, err_result)
                self.batch_error.emit(filename, str(e))

        self.batch_finished.emit(summary)
