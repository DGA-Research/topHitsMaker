# Top Hits Maker

This Streamlit app reformats Word documents that use Heading 2-Heading 5 styles into a concise "Key Points" summary with consistent formatting.

## Virtual Access
- App can be accessed virtually via: https://a685m2sxhb9zgrmdr62bnf.streamlit.app/

## Local Prerequisites (optional)
- Python 3.9 or newer
- The packages listed in `requirements.txt`

## Local Installation (optional)
1. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv .venv
   .\.venv\Scripts\activate    # Windows
   source .venv/bin/activate   # macOS/Linux
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Locally Run the App (optional)
Launch the Streamlit server from the project root:
```bash
streamlit run app.py
```
Streamlit will open a browser tab (or provide a local URL) for the app.

## Using the App
- Upload a `.docx` file that uses Word’s Heading 2, Heading 3, Heading 4, and Heading 5 styles.
- Choose whether Heading 4 bullets should always end with a period.
- Download the generated document (`*_OUTPUT.docx`) once processing finishes.

## Notes
- The output document uses Arial 10pt, single spacing, and narrow margins.
- Content not styled as Heading 2-Heading 5 is ignored on purpose.
- Example input and output files are included in the repo for quick testing.
