Using templates (DOCX) and PDF parsing

Environment variables
- Set `OFICIO_TEMPLATES_DIR` to the folder containing your DOCX templates.
- Optionally set `OFICIO_TEMPLATE` to a specific template file. It can be:
  - an absolute/relative path to a `.docx`, or
  - just the filename inside `OFICIO_TEMPLATES_DIR`.

Resolution rules
- If `OFICIO_TEMPLATE` points to an existing file, it is used.
- Else, if `OFICIO_TEMPLATES_DIR` exists, the first `.docx` in that folder is used.
- Else, it falls back to `templates/oficio_modelo.docx` if present.

Placeholders in DOCX
- The generator replaces placeholders like `{{NUM_PROCESSO}}`, `{{DATA}}`.
- It also fills (when found in the PDF): `{{ASSUNTO}}`, `{{INTERESSADO}}`, `{{REQUERENTE}}`, `{{DATA_DOCUMENTO}}`, and `{{EXTRATO}}`.

Example (.env)
OFICIO_TEMPLATES_DIR=C:\Users\20386\OneDrive - tcm.sp.gov.br\Oficios\oficio_automation\Modelos Oficios
OFICIO_TEMPLATE=meu_modelo.docx

