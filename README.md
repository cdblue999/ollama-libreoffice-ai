# Ollama LibreOffice AI

Local AI document assistant for LibreOffice (Writer, Calc, Impress) — powered by Ollama.

## Description

Ollama LibreOffice AI is a LibreOffice Basic macro that adds AI-powered document assistance using locally-running Ollama models. It is a LibreOffice adaptation of [Ollama Outlook AI](https://github.com/cdblue999/ollama-outlook-ai).

No cloud dependency. No API keys. No data leaves your machine.

## Features

- **Fully local and offline** — works without internet
- **No API key required** — Ollama runs on your own computer
- **Compatible with any Ollama model** (llama3.2, mistral, gemma, phi, etc.)
- **Privacy-preserving** — document content never leaves your machine
- **Supports Writer, Calc, and Impress** documents
- **Configurable system prompt** for custom AI behavior
- **Model selection** dropdown populated from your installed Ollama models
- **Adjustable timeout** for slower models
- **Windows only** (uses COM for HTTP, registry, shell)

## Prerequisites

1. **Ollama** installed and running — download from [ollama.com](https://ollama.com)
2. **At least one model pulled** — run `ollama pull llama3.2:3b` (2GB, runs well on CPU)
3. **LibreOffice 7.x or later** (any app: Writer, Calc, Impress)
4. **Windows** (uses WinHttp, WScript.Shell, and PowerShell via COM)

## Installation

### Option A: Import .xba (easiest)

1. Open **LibreOffice Writer** (or Calc, Impress)
2. **Tools → Macros → Organize Macros → LibreOffice Basic**
3. Click **Edit** to open the Basic IDE
4. In the Basic IDE: **File → Import Basic Module...** (or right-click in module tree → Import Basic Module...)
5. Select `OllamaAI.xba` from this repo
6. **Ctrl+S** to save
7. Close the Basic IDE
8. **Tools → Macros → Run Macro → Ollama_Initialize → Run** (only needed once)

### Option B: Copy-paste .bas

1. Open `OllamaAI.bas` in Notepad → **Ctrl+A → Ctrl+C**
2. LibreOffice: **Tools → Macros → Organize Macros → LibreOffice Basic → Edit**
3. In the left tree, right-click **Standard → Insert → Module**
4. Name it: **OllamaAI**
5. **Ctrl+V → Ctrl+S**
6. **Tools → Macros → Run Macro → Ollama_Initialize → Run** (only needed once)

## Configuration

To configure the addon:

1. **Tools → Macros → Run Macro → Ollama_ShowConfigurationForm → Run**
2. A settings window appears where you can:
   - **Select Model**: Choose from your installed Ollama models (fetched automatically)
   - **Custom Prompt**: Edit the system prompt to tailor AI responses
   - **Timeout**: Adjust API timeout (default 120s, increase for slower models)

All settings are saved in Windows Registry under `HKCU\Software\OllamaLibreOfficeAI`.

## Usage

1. Open a **Writer**, **Calc**, or **Impress** document
2. **Tools → Macros → Run Macro → Ollama_ProcessDocument → Run**
3. The macro reads your document content and sends it to Ollama
4. The AI response is inserted at the cursor position (Writer) or into the active cell (Calc)

### Document type behavior

| Document Type | Content Extracted | Response Inserted |
|---------------|------------------|-------------------|
| Writer | Full document text | At cursor position |
| Calc | All sheets, used cells | Active cell or A1 of first sheet |
| Impress | All slides, text from text shapes | Selected text shape (select one first) |

## How It Works

1. The macro reads the current document's content using LibreOffice's UNO API
2. It sends the content to Ollama's OpenAI-compatible API at `http://localhost:11434/v1/chat/completions`
3. Ollama processes the request using your selected model
4. The AI-generated response is inserted into your document

For screenshot analysis (`Ollama_AnalyzeScreenshot`), a PowerShell file dialog picks the image, converts it to base64, and sends it as a `data:` URI to Ollama's vision API.

All communication is via HTTP to `localhost` — no data travels over the network.

## Available Macros

| Macro | Description |
|-------|-------------|
| `Ollama_Initialize` | Create default registry settings (run once after import) |
| `Ollama_ProcessDocument` | Send active document to Ollama AI and insert response |
| `Ollama_ShowConfigurationForm` | Open settings window (model, prompt, timeout) |
| `Ollama_ShowAboutForm` | Show version and info |
| `Ollama_IsOllamaRunning` | Check if Ollama server is reachable (returns Boolean) |
| `Ollama_RefreshModelsList` | Refresh the cached list of available models |
| `Ollama_AnalyzeScreenshot` | Pick an error screenshot/image file, send to a vision model, and insert analysis |

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Ollama is not running" | Open a terminal and run `ollama list` to verify. Start Ollama from Start Menu if not running. |
| No models in dropdown | Run `ollama pull llama3.2:3b` in terminal, then open Settings again. |
| Request times out | Increase timeout in Settings (120s+). Smaller models respond faster. |
| "Unsupported document type" | Only Writer (.odt, .docx), Calc (.ods, .xlsx), and Impress (.odp, .pptx) are supported. |
| Screenshot analysis returns empty/error | The selected model may not support vision. Switch to a multimodal model like `llama3.2-vision` or `llava` in Settings. |
| Settings window not showing | PowerShell must be available on your system (built into Windows). |
| VBA-style `_` line continuations | LibreOffice Basic does not support them — this module uses single-line format. |

## Technical Details

- **File to import**: `OllamaAI.xba` (XML Basic module for Macro Organizer) or `OllamaAI.bas` (plain text for copy-paste)
- **Language**: StarBasic / LibreOffice Basic
- **API Format**: OpenAI-compatible chat completions endpoint
- **Document Access**: LibreOffice UNO API (`ThisComponent`, `Text`, `Sheets`, `Slides`)
- **HTTP Client**: WinHttp.WinHttpRequest.5.1 (COM)
- **Configuration Storage**: Windows Registry (`HKCU\Software\OllamaLibreOfficeAI`)
- **UI Generation**: PowerShell script (invoked from Basic)
- **No external dependencies** — uses built-in Windows and LibreOffice components only

## Files

| File | What |
|------|------|
| `OllamaAI.xba` | Importable module — Basic IDE → File → Import Basic Module... |
| `OllamaAI.bas` | Plain text source — open, copy, paste manually |
| `README.md` | This file |

## Data Flow

```
LibreOffice (Basic)  --HTTP--> localhost:11434 -- Ollama Model
       ^                         |
       |                         v
       +------ JSON Response ----+
```

## Privacy

This addon is fully local:
- **No data** is sent to any external server
- **No API keys** are stored or transmitted
- **No telemetry**, analytics, or tracking
- **No internet connection required** after Ollama is installed
- All document processing happens on your machine

## License

MIT License

Copyright (c) 2026 cdblue999

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
