# 📜 Dynamic Certificate Generator

A **frontend-only web application** to generate personalized certificates from a DOCX template and an Excel file. Supports dynamic placeholders for parameters like `name`, `rank`, `category`, etc., and allows batch downloading of certificates.

---

## Features

- Add **custom parameters** dynamically.
- Upload **DOCX certificate template** with placeholders.
- Upload **Excel file** with corresponding data.
- Generate **one certificate per row** in Excel.
- **Dark mode** UI with responsive design.
- **Modal guide** on how to use the application.

---

## How to Use

1. Click **"How to Use?"** button on the top-right for instructions.
2. **Step 1:** Add parameters (e.g., `name`, `rank`, `category`) → Click **Add**.
3. After adding all parameters → Click **Completed**.
4. **Step 2:** Upload your DOCX template with the same placeholders.
5. Upload your Excel file with column names matching parameters.
6. Click **Generate Certificates** → it will download one certificate per row.  
   💡 Tip: Your browser may ask permission to allow multiple downloads. Click **Allow** to download all certificates in one go.

---

## DOs ✅

- Use **DOCX format only**.
- Placeholders must match exactly: `{name}`, `{rank}`, `{category}` or `{{name}}`, `{{rank}}`, `{{category}}`.
- Excel column names must match placeholders (lowercase).
- Keep Word template simple (no text boxes or shapes on top).

---

## DON'Ts ❌

- Do not put text on top of images inside Word.
- Do not rename Excel columns with extra spaces.
- Do not use merged cells in Excel.
- Do not use `.doc` or `.csv` files.

---

## Common Problems & Fixes

- **Missing name → skipping row:** Fill blank cells in Excel.  
- **TemplateError / Multi error:** Check your Word placeholders.  
- **Docxtemplater not available:** Ensure scripts are loaded correctly.

---

## Installation

1. Clone or download this repository.
2. Open `index.html` in any modern browser (Chrome, Edge, Firefox).  
   No backend or server required — fully frontend solution.

---

## Dependencies

- [PizZip](https://github.com/open-xml-templating/pizzip) – for reading DOCX files
- [Docxtemplater](https://github.com/open-xml-templating/docxtemplater) – for template rendering
- [SheetJS (XLSX)](https://github.com/SheetJS/sheetjs) – for Excel parsing

All libraries are included via CDN in `index.html`.

---

## Folder Structure

```bash
/dynamic-certificate-generator
│
├── index.html # Main HTML page
├── style.css # CSS styling
├── script.js # JavaScript logic
├── sample.docx # Example certificate template
├── sample.xlsx # Example Excel data
└── README.md # Project documentation

```

---

## License

This project is open-source and free to use under the **MIT License**.


