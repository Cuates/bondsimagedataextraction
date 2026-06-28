# US Savings Bond Data Extractor

This script automates the extraction of structured data from photos of US Savings Bonds using Google's **Gemini 3.1 Flash Lite** AI model. It processes a folder of images (including skewed or upside-down iPhone photos) and exports the details into a clean Excel (`.xlsx`) spreadsheet.

## Features

- **High-Accuracy Extraction:** Extracts Serial Number, Series, Denomination, Issue Date, Owner Names, Co-Owner, Address, and Notes.
- **Multimodal AI:** Utilizes Gemini 3.1 Flash Lite to automatically interpret skewed, rotated, or upside-down images without needing complex computer vision pre-processing.
- **Strict JSON Parsing:** Uses Pydantic to enforce a strict structured response schema, guaranteeing clean spreadsheet columns.
- **Predictable Sorting:** Images are automatically processed in alphabetical order, and the final exported Excel rows are also sorted alphabetically by filename for easy auditing.
- **Dual Processing Modes:**
  - **Sequential (Free Tier):** Processes images one at a time with a mandatory 4.1-second delay, ensuring compliance with Google's free tier rate limits (15 Requests Per Minute).
  - **Concurrent (Paid Tier):** Uses a `ThreadPoolExecutor` to blast through 15 images simultaneously for massive speed upgrades when unconstrained by free quotas.
- **Auto-Resume & Preservation:** If the script halts due to quota exhaustion, it preserves already processed images. When run again, it automatically detects the existing Excel file and skips images it has already completed.
- **Robust Error Handling:** Wraps requests in exponential backoff retries to handle temporary server spikes.

---

## Prerequisites

1. **Python 3.10+** installed on your system.
2. A **Google Gemini API Key**. You can get one from [Google AI Studio](https://aistudio.google.com/app/apikey).

## Installation

Open your terminal and install the required dependencies:

```bash
pip install google-genai pydantic pandas openpyxl tqdm
```
*(Note: `concurrent.futures`, `json`, `os`, and `mimetypes` are built into standard Python).*

## Configuration

Before running the script, you must expose your Gemini API key to your environment.

**If you are using PowerShell (Windows default):**
```powershell
$env:GEMINI_API_KEY="your_api_key_here"
```

**If you are using Command Prompt (cmd):**
```cmd
set GEMINI_API_KEY=your_api_key_here
```

---

## Usage

Run the script from your terminal:

```bash
python extract_bonds.py
```

The script is interactive and will prompt you for three things:

1. **API Tier Selection:**
   - `Are you using the Free Tier API? (y/n)`
   - Type `y` to use the slow, safe sequential mode to avoid getting rate limited.
   - Type `n` if you have billing enabled on Google Cloud and want lightning-fast concurrent processing.

2. **Input Directory:**
   - Provide the absolute or relative path to the folder containing your bond images (supports `.png`, `.jpg`, `.jpeg`, and `.heic`).

3. **Output Filename:**
   - Provide the name for your Excel file (e.g., `bonds_data.xlsx`).

### Resuming from an Interruption

If you hit your daily/minute quota limits (Error 429), the script will save what it successfully processed before exiting.
To resume:
1. Run the script again.
2. Provide the **exact same Output Filename**.
3. The script will read the existing Excel file, skip all successfully processed images, and continue precisely where it left off!

---

## Troubleshooting

- **`No module named ...` or `Unable to import ...`**
  Ensure you ran the `pip install` command in the same Python environment that you are using to run the script.
- **`Error: APIError - 429 RESOURCE_EXHAUSTED`**
  You have hit your Google API quota. If you are on the free tier, wait a few minutes (or until tomorrow if you hit the daily limit) and run the script again to resume.
- **`Failed to initialize Gemini Client.`**
  Your `GEMINI_API_KEY` was not found. Make sure you exported the environment variable exactly as shown in the Configuration section before running the script.

---

## Cleanup (Optional)

If you are completely finished with this script and want to remove the installed tools and your API key from your system, you can run the following:

**Uninstall the Python dependencies:**
```bash
pip uninstall -y google-genai pydantic pandas openpyxl tqdm
```

**Remove the API Key from your current session:**
- **PowerShell:**
  ```powershell
  $env:GEMINI_API_KEY=""
  ```
- **Command Prompt (cmd):**
  ```cmd
  set GEMINI_API_KEY=
  ```
