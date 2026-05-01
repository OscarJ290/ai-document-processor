# 🤖 AI Document Processor — Portfolio Project

Extracts structured data from **any document** (PDFs, images, invoices, contracts) using the **Claude API**. Outputs clean JSON and Excel reports automatically. Zero templates required.

Built by **Oscar J. Villa García** · [ojviga@gmail.com](mailto:ojviga@gmail.com)

---

## 🚀 Quick Start

```bash
git clone https://github.com/oscarjvilla290/ai-document-processor.git
cd ai-document-processor
pip install -r requirements.txt
echo "ANTHROPIC_API_KEY=your_key_here" > .env

# Single file
python src/processor.py invoice.pdf --hint "extract invoice number, total, date"

# Batch folder
python src/processor.py ./documents/ --label "batch_invoices"
```

---

## ✨ Features

- Supports PDFs, images (PNG/JPG/WEBP), and text files
- Claude AI understands context — no regex, no templates
- Exports to JSON + styled Excel automatically
- Batch processing for entire folders
- Optional extraction hints for focused output

---

## 💡 Use Cases

Invoice processing · Contract extraction · Form digitization · ERP data entry automation

---

## 🛠️ Tech Stack

**Python 3.11+** · **Claude API** · `requests` · `pandas` · `openpyxl`

---

## 📄 License

MIT
