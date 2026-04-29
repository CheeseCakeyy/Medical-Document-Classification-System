# MediSort — Medical Document Classifier
### A PySide6 Desktop Application for Automatic Medical Document Organization

---

## Project Structure

```
medical_doc_classifier/
│
├── main.py                   # GUI application (PySide6)
├── backend.py                # OCR, model inference, file sorting logic
├── requirements.txt          # Python dependencies
├── generate_dummy_model.py   # Helper: creates test model.pkl / vectorizer.pkl
│
├── model.pkl                 # ← Your trained scikit-learn model (place here)
├── vectorizer.pkl            # ← Your trained TF-IDF vectorizer (place here)
│
└── classified_documents/     # Auto-created; sorted output folder
    ├── prescription/
    ├── lab_report/
    ├── discharge_summary/
    ├── medical_bill/
    ├── insurance_claim/
    ├── referral_letter/
    └── other/
```

---

## Prerequisites

### 1. Python
Requires **Python 3.9+**.

```bash
python --version
```

### 2. Tesseract OCR Engine
Tesseract must be installed system-wide (separate from the Python package).

**Windows:**
1. Download installer from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install (e.g. to `C:\Program Files\Tesseract-OCR\`)
3. Add to system PATH, OR set the path in your script:
   ```python
   # Add at the top of backend.py if needed:
   import pytesseract
   pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
   ```

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

Verify installation:
```bash
tesseract --version
```

---

## Installation

### Step 1: Create a virtual environment (recommended)

```bash
python -m venv venv

# Activate:
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate
```

### Step 2: Install Python dependencies

```bash
pip install -r requirements.txt
```

---

## Setup Your Model Files

### Option A — Use your real trained model (production)

Place your trained files directly in the project folder:
```
model.pkl        ← your scikit-learn classifier
vectorizer.pkl   ← your TF-IDF vectorizer
```

Both must be saved with `pickle`. The model must:
- Have a `predict(X)` method
- Optionally have `predict_proba(X)` for confidence scores
- Return predictions matching one of the 7 category names:
  `prescription`, `lab_report`, `discharge_summary`,
  `medical_bill`, `insurance_claim`, `referral_letter`, `other`

### Option B — Generate dummy test models (development/testing)

If you want to test the UI before integrating your real model:

```bash
python generate_dummy_model.py
```

This creates minimal `model.pkl` and `vectorizer.pkl` for UI testing only.

---

## Running the Application

```bash
python main.py
```

---

## How to Use

1. **Launch** the application — the model loads automatically at startup.
2. **Upload a document** — click "Upload Document" or drag & drop a file into the preview panel.
   - Supported formats: `.pdf`, `.png`, `.jpg`, `.jpeg`, `.tiff`, `.bmp`
3. **Classify** — click "Classify Document" to run OCR + classification.
4. **View result** — the predicted category appears in the result panel with a confidence score.
5. **View OCR text** — click "View OCR Text" to inspect what was extracted.
6. **Open output folder** — click "Open Output Folder" to browse classified files.

---

## Application Workflow (Internal)

```
User uploads file
       │
       ▼
OCREngine.extract_text()
  ├── PDF  → fitz direct extraction → OCR fallback if no text layer
  └── Image → OpenCV preprocessing → pytesseract
       │
       ▼
ModelManager.predict(text)
  ├── vectorizer.transform(text)   ← TF-IDF features
  └── model.predict(features)      ← category label
       │
       ▼
FileOrganizer.save_document()
  └── classified_documents/<category>/<filename>
       │
       ▼
GUI displays result
```

---

## Customization

### Change output folder
In `main.py`, pass a custom path when creating `DocumentClassifier`:
```python
self.classifier = DocumentClassifier(output_dir="/path/to/my/output")
```

### Change model/vectorizer paths
```python
self.classifier = DocumentClassifier(
    model_path="models/my_model.pkl",
    vectorizer_path="models/my_vectorizer.pkl"
)
```

### Add/remove categories
Edit the `CATEGORIES` list in `backend.py` and update `CATEGORY_COLORS` and `CATEGORY_ICONS` in `main.py`.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `TesseractNotFoundError` | Install Tesseract and add to PATH (see Prerequisites) |
| `FileNotFoundError: model.pkl` | Place your model files in the project folder or run `generate_dummy_model.py` |
| OCR returns empty text | Check file is readable; try a cleaner scan with higher DPI |
| PDF shows blank OCR | PDF may be image-based; OCR fallback will run automatically |
| Poor classification accuracy | This is a model quality issue — retrain or improve your model |
| GUI looks blurry on Windows | Set system DPI scaling to 100% or enable DPI awareness in Windows settings |

---

## Dependencies

| Package | Purpose |
|---|---|
| `PySide6` | GUI framework |
| `scikit-learn` | ML model inference |
| `pytesseract` | Python wrapper for Tesseract OCR |
| `opencv-python` | Image preprocessing |
| `Pillow` | Image loading fallback |
| `PyMuPDF` | PDF text extraction + rendering |
| `numpy` | Array operations |
