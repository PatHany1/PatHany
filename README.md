# Resume Tailor

`resume_tailor.py` customizes a base resume for a given company and role using an LLM and outputs a clean `.docx`.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### OpenAI backend

```bash
export OPENAI_API_KEY=sk-...
python resume_tailor.py \
  --base-resume ./1-24-25.docx \
  --company "WSP (Parsons Brinckerhoff)" \
  --role "Mechanical Engineer — Buildings/HVAC" \
  --job-desc ./wsp_mech_jd.txt \
  --backend openai --model gpt-5 \
  --out "Gavin_Joyce_Resume_WSP.docx"
```

### Ollama backend

```bash
python resume_tailor.py \
  --base-resume ./1-24-25.docx \
  --company "Acme MEP" \
  --role "Junior Mechanical Engineer" \
  --job-desc-text "Acme seeks a junior ME to support HVAC and plumbing design using Revit..." \
  --backend ollama --model llama3 \
  --out "Gavin_Joyce_Resume_Acme.docx"
```

## Exit codes

* `0` – success
* non‑zero – failure; check stdout for details. If model output isn't valid JSON, a `*_RAW.txt` file will be written alongside the requested output for inspection.

