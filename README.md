# 📝 word-to-pdf-multilang

Simple CLI tool to convert `.docx` files to PDF using Microsoft Word, with multilingual error and status messages.  
**Now supports:** English, Spanish, French, Italian and Russian. 🇬🇧 🇪🇸 🇫🇷 🇮🇹 🇷🇺

> ⚠️ This tool requires Microsoft Word — works only on **Windows** and **macOS**.

---

## ✨ Features

- Convert a single `.docx` file or an entire directory of `.docx` files to PDF
- Multilingual messages: `en`, `es`, `fr`, `it`, `ru`
- Optional output directory via `--saveat`
- Language override via `--lang`
- Simple and minimal CLI interface

---

## 🧪 Requirements

- Python 3.6+
- Microsoft Word installed (Windows/macOS only)
- [`docx2pdf`](https://pypi.org/project/docx2pdf/)

---

## 📦 Requirement Installation

```bash
pip install docx2pdf
```

---

## 🚀 Usage

```bash
python word_to_pdf.py <path_to_docx_or_directory> [--saveat <output_directory>] [--lang <language_code>]
```

### 🔹 Examples

Convert a single file:

```bash
python word_to_pdf.py myfile.docx
```

Convert a file with an output folder:

```bash
python word_to_pdf.py myfile.docx --saveat ./converted
```

Convert a directory of `.docx` files with Russian messages:

```bash
python word_to_pdf.py ./documents --lang ru
```

---

## 🌍 Available Language Codes

| Code | Language    |
|------|-------------|
| en   | English     |
| es   | Spanish     |
| fr   | French      |
| it   | Italian     |
| ru   | Russian     |

If no `--lang` is provided, the system language will be used if supported, otherwise defaults to English language.

---

## 🛡 License

This project is licensed under the [MIT License](LICENSE).

