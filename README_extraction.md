# AISESA — AI Extraction Tool (local)

Upload PDFs (one or many) → AI extracts all 58 fields with source quotes →
verify each in a side-by-side view with the PDF → export to relational Excel.

## Setup
```bash
pip install -r requirements_extraction.txt
export GEMINI_API_KEY=your_key        # or paste in the sidebar
streamlit run extraction_app.py
```

## The three tabs

**📤 Upload & extract** — Drag in one PDF or 50. Each is extracted and saved as
a **draft** (not yet verified). The whole batch runs sequentially; failures are
recorded with their error message so you can see what went wrong.

**📝 Review drafts** — Pick any draft to open the verification view. You see the
PDF on the right, the extracted text in a second tab on the right (searchable),
and the form with every field on the left, each with its AI-extracted value
and the source quote that justifies it.

**✅ Verified** — Your saved studies. Bulk-delete with checkboxes if needed.

## Key features

- **Batch upload**: select N PDFs at once, leave it running.
- **Drafts persist**: close & reopen the app, drafts are still there.
- **Failures kept**: scanned PDFs / timeouts are recorded with their error.
- **PDF viewer + extracted text** side-by-side with the form.
- **Power pools computed** from countries — never guessed.
- **Synthesis upgrade**: study_objective / key_result on Pro while rest on Flash.

## Costs (Gemini 2.5)
- Flash only: ~$3–4 for 1400 articles.
- Flash + synthesis upgrade: ~$17 for 1400 (recommended sweet spot).
- Pro for everything: ~$40 for 1400.

## Workflow
- **Gold-standard test**: upload 10–15 of your hand-extracted studies, batch
  them, then review to compare AI vs your answers.
- **Production**: 50 at a time, ~15 min batch, then verify.
- **PDF cache**: kept in _pdf_cache/. Reset wipes everything.
