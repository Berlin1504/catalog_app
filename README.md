# Product Catalog App (Flask)

A small web app to build a veterinary product catalog in DOCX format.
Add products with text fields and images; download the formatted catalog
that matches the original template (green banners + yellow 5-column tables).

## Features

- 🏷️ Two configurable line banners ("الخط الأول" / "الخط الثاني")
- ➕ Add products with name, composition, indications, features
- 🖼️ Upload a product image per product (appears in the العبوة column)
- 📋 Live list of added products with delete option
- ⬇️ One-click DOCX download with the exact template layout
- 🇸🇦 Full Arabic RTL support
- 💾 Data persisted to `data.json`; images saved to `uploads/`

## Setup

```bash
cd catalog_app
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Then open <http://localhost:5000> in your browser.

## Usage

1. **Set banner names** — Top of the page. Edit الخط الأول / الخط الثاني and save.
2. **Add a product** — Fill in name, composition (one line per item), indications, features, optionally upload an image, choose which banner to attach the product to, then click "إضافة المنتج".
3. **Manage products** — Each added product appears in the list. Delete with the red button.
4. **Generate** — Click "تحميل الكتالوج (DOCX)" to download the catalog.

## Output

The DOCX will have, per non-empty banner:

```
[ GREEN BANNER: line name ]

[YELLOW HEADER ROW]
| المميزات | دواعي الاستخدام | التركيب | العبوة (image) | اسم المنتج (yellow) |
```

Followed by one row per product. The page is A4 landscape.

## Files

```
catalog_app/
├── app.py                  # Flask app + DOCX generator
├── templates/
│   └── index.html          # Single-page form
├── static/
│   └── style.css           # RTL-friendly CSS
├── uploads/                # Saved product images (created at runtime)
├── output/                 # Generated DOCX files (created at runtime)
├── data.json               # Persisted form state (created at runtime)
└── requirements.txt
```

## Notes

- All product data is stored locally in `data.json`; nothing is sent to a server.
- Images are stored in `uploads/` and referenced from the generated DOCX.
- The "Clear all" button wipes `data.json` and uploaded images.
- The app runs in debug mode by default — disable for production.