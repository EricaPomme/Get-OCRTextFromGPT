# Get-OCRTextFromGPT

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
  - [Context menu integration](#context-menu-integration)
- [Usage](#usage)
  - [Basic usage](#basic-usage)
  - [Multi-page documents](#multi-page-documents)
  - [Saving to a file](#saving-to-a-file)
  - [Choosing a model](#choosing-a-model)
  - [Verbose output](#verbose-output)
- [Parameters](#parameters)
- [Output format](#output-format)
  - [Text](#text)
  - [Tables](#tables)
  - [Embedded images and charts](#embedded-images-and-charts)
  - [Non-English text](#non-english-text)
- [Chat transcript mode](#chat-transcript-mode)
- [Privacy and data retention](#privacy-and-data-retention)
  - [What the script does](#what-the-script-does)
  - [What cannot be eliminated via API parameters](#what-cannot-be-eliminated-via-api-parameters)
- [Known limitations](#known-limitations)

---

## Overview

`Get-OCRTextFromGPT.ps1` converts document images and chat screenshots to
markdown by sending them to an OpenAI vision model. Multiple images are
treated as sequential pages of one document (or sequential scrolls of one
conversation) -- context carries forward so the model can handle elements
that span page boundaries.

Output follows these conventions:

- Text is transcribed verbatim.
- Embedded images and photographs are described as `[Image: ...]`.
- Charts and graphs are described as `[Chart: ...]`.
- Tables are rendered as GitHub-Flavored Markdown tables.
- Blocks of text wholly in a non-English language are followed by an inline
  machine translation blockquote.
- Chat screenshots are auto-detected and formatted as transcripts with
  speaker names and resolved timestamps (see [Chat transcript mode](#chat-transcript-mode)).

---

## Prerequisites

- **PowerShell 5.1 or later.** Tested on Windows PowerShell 5.1 and
  PowerShell 7+.
- **.NET Framework 4.x** (included with PowerShell 5.1 on Windows). Required
  for `System.Drawing`, which is used to strip EXIF metadata from images.
- **An OpenAI API key** with access to a vision-capable model (`gpt-5.5` by
  default). Set the `OPENAI_API_KEY` environment variable or pass the key via
  `-ApiKey`.

---

## Installation

Copy `Get-OCRTextFromGPT.ps1` to a location of your choice. No module
installation or additional dependencies are required.

If PowerShell execution policy prevents running scripts, either unblock the
file:

```powershell
Unblock-File .\Get-OCRTextFromGPT.ps1
```

or set an appropriate execution policy for your session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### Context menu integration

`Add-ContextMenuItems.cmd` and `Remove-ContextMenuItems.cmd` are included in
the repository alongside the script. Running `Add-ContextMenuItems.cmd` (as a
normal user -- no elevation required) registers two entries under
`HKCU\Software\Classes\SystemFileAssociations\image\shell`, which adds them to
the right-click menu for all image file types recognized by Windows (PNG, JPEG,
GIF, WebP, etc.):

| Menu item           | Equivalent command-line flags |
| ------------------- | ----------------------------- |
| **OCR**             | `-ToClipboard`                |
| **OCR (Chat Mode)** | `-ToClipboard -ChatMode`      |

Both entries invoke `powershell.exe` (Windows PowerShell 5.1) with `-NoProfile`
to avoid profile side effects. The script path is resolved from the location of
the `.cmd` file at install time, so no manual path editing is required as long
as the files stay in the same folder.

To remove the entries, run `Remove-ContextMenuItems.cmd`.

---

## Usage

### Basic usage

```powershell
.\Get-OCRTextFromGPT.ps1 scan.png
```

Converts a single image and writes the markdown to stdout.

### Multi-page documents

```powershell
.\Get-OCRTextFromGPT.ps1 page1.jpg, page2.jpg, page3.jpg
```

Pages are processed in the order given. Conversation context carries forward
so the model does not repeat the document title or headings on subsequent pages.

To process all PNG files in a folder, sorted by name:

```powershell
$pages = Get-ChildItem *.png | Sort-Object Name |
         Select-Object -ExpandProperty FullName
.\Get-OCRTextFromGPT.ps1 -Images $pages
```

### Saving to a file

```powershell
.\Get-OCRTextFromGPT.ps1 -Images scan.png -OutputPath output.md
```

The markdown is written to both stdout and `output.md`. The file is encoded
in UTF-8 without BOM for cross-tool compatibility.

### Choosing a model

```powershell
.\Get-OCRTextFromGPT.ps1 scan.png -Model gpt-4o -MaxTokens 2048
```

Supported model families:

| Family  | Example IDs              | Notes                                              |
| ------- | ------------------------ | -------------------------------------------------- |
| GPT-5.5 | `gpt-5.5`, `gpt-5.5-...` | Default. `detail: original` (6000px, 10000 tiles). |
| GPT-5.x | `gpt-5.4`, `gpt-5-mini`  | `detail: high`. No temperature/penalty params.     |
| GPT-4o  | `gpt-4o`, `gpt-4.1`      | `detail: high`. Supports temperature.              |

The script automatically detects the model family and sends the correct
parameter set for each (`max_completion_tokens` vs `max_tokens`, presence or
absence of `temperature`).

### Verbose output

```powershell
.\Get-OCRTextFromGPT.ps1 scan.png -Verbose
```

Prints per-image progress and model configuration to the verbose stream
without affecting the markdown on stdout.

---

## Parameters

| Parameter      | Type       | Default          | Description                                                                                           |
| -------------- | ---------- | ---------------- | ----------------------------------------------------------------------------------------------------- |
| `-Images`      | `string[]` | _(required)_     | One or more image paths (PNG, JPEG, GIF, WebP). Positional (no name needed in position 0).            |
| `-OutputPath`  | `string`   | _(none)_         | If set, also writes the markdown to this file (UTF-8, no BOM).                                        |
| `-ApiKey`      | `string`   | `OPENAI_API_KEY` | OpenAI API key. Overrides the environment variable.                                                   |
| `-Model`       | `string`   | `gpt-5.5`        | Vision-capable chat completions model. See model table above.                                         |
| `-MaxTokens`   | `int`      | `4096`           | Maximum tokens per page response. Increase for very dense documents, decrease to reduce cost.         |
| `-ChatMode`    | `switch`   | _(off)_          | Forces chat transcript mode. Skips auto-detection. See [Chat transcript mode](#chat-transcript-mode). |
| `-ToClipboard` | `switch`   | _(off)_          | Copies the final markdown to the system clipboard in addition to stdout.                              |

---

## Output format

### Text

All text is transcribed verbatim -- capitalization, punctuation, and spacing
are preserved. Illegible characters are marked `[?]`. Longer unreadable
passages are marked `[...]`.

### Tables

Tables in the source document are converted to GitHub-Flavored Markdown tables:

```markdown
| Column A | Column B | Column C |
| -------- | -------- | -------- |
| value 1  | value 2  | value 3  |
```

### Embedded images and charts

Non-text elements are described inline:

```text
[Image: Company logo. Shield shape. Dark blue background. White text.]
[Chart: Line graph. X-axis: Jan-Dec 2025. Y-axis: units sold. Single line. Peaks in March and November.]
```

### Non-English text

When a complete paragraph or block is written entirely in a non-English
language, it appears verbatim followed immediately by a translation blockquote:

```markdown
Sehr geehrte Damen und Herren, wir mochten Sie uber die Anderungen
in unserer Datenschutzrichtlinie informieren.

> \_[Machine translation from German: Dear Sir or Madam, we would like
>
> > to inform you about the changes to our privacy policy.]\_
```

Single words, loan words, proper nouns, and short foreign-language expressions
embedded naturally within English text are left untranslated.

---

## Chat transcript mode

Screenshots of chat or messaging applications (Microsoft Teams, Slack, Discord,
iMessage, WhatsApp, and similar) are automatically detected and formatted as
transcripts. No flags are required -- the script makes a brief classification
API call on the first image and switches modes if a chat UI is detected.

To skip auto-detection and force transcript mode:

```powershell
.\Get-OCRTextFromGPT.ps1 -Images teams-screenshot.png -ChatMode
```

To process multiple sequential scrolls of the same conversation:

```powershell
.\Get-OCRTextFromGPT.ps1 teams-p1.png, teams-p2.png, teams-p3.png -OutputPath transcript.md
```

### Transcript format

Each message is rendered as:

```markdown
**Sender Name** (YYYY-MM-DD H:MM am/pm)

Message content here.
```

### Timestamp resolution

Relative timestamps in chat UIs are resolved to absolute dates using the
current system date and time at the moment the script runs. The priority
order for inference is:

1. **Explicit date in the screenshot** -- always wins (e.g. a "Monday, May 26"
   separator bar).
2. **Chronological order** -- messages are in ascending time order, which
   narrows the range.
3. **Time-of-day comparison** -- a bare time (e.g. "5:04 pm") with no date
   context is compared against the current clock time. If the time shown is
   earlier than or equal to now, it belongs to today. If it is later than now
   (i.e. that hour has not yet occurred today), it belongs to yesterday.

Examples (assuming the script runs on 2026-05-27 at 8:40 am):

| Screenshot shows | Reasoning                       | Resolves to        |
| ---------------- | ------------------------------- | ------------------ |
| `8:50 am`        | 8:50 am > 8:40 am (current)     | 2026-05-26 8:50 am |
| `8:30 am`        | 8:30 am <= 8:40 am (current)    | 2026-05-27 8:30 am |
| `5:04 pm`        | 5:04 pm > 8:40 am (current)     | 2026-05-26 5:04 pm |
| `Yesterday`      | Explicit relative day           | 2026-05-26         |
| `Monday`         | Most recent Monday on or before | 2026-05-25         |
| `May 5`          | Month/day in current year       | 2026-05-05         |

### What is and is not transcribed

| Element                             | Handling                         |
| ----------------------------------- | -------------------------------- |
| Message text                        | Transcribed verbatim             |
| Shared photos, document previews    | `[Image: description]`           |
| "see more" / truncated messages     | Visible text + truncation note   |
| Profile picture thumbnails          | Skipped                          |
| Emoji reactions, read receipts      | Reactions included if meaningful |
| App chrome (sidebar, toolbar, etc.) | Skipped                          |

### Example output

```markdown
**Alice Example** (2026-05-26 3:47 pm)

hey are you joining the call later

**Bob Example** (2026-05-26 3:48 pm)

yeah, be there in 5

[message continues -- truncated in screenshot]
```

---

## Privacy and data retention

### What the script does

Before transmitting any image, the script re-encodes it as PNG via
`System.Drawing`. This strips all EXIF metadata client-side -- GPS coordinates,
device identifiers, software tags, and timestamps are removed before the bytes
leave your machine.

Each API request includes:

- `store: false` -- prevents the completion from being stored in OpenAI's
  Stored Completions API (used for distillation and evals). This is the default
  but is set explicitly for clarity.
- `X-OpenAI-No-Store: true` header -- requests Zero Data Retention behavior.
  Honored only for enterprise accounts with a ZDR agreement.
- `X-Stainless-*: private` headers -- suppresses SDK environment telemetry
  (OS, architecture, runtime version) from the request.
- No `user`, `safety_identifier`, `metadata`, or `prompt_cache_key` fields --
  avoids attaching identity or caching hints to the request.

### What cannot be eliminated via API parameters

OpenAI retains API inputs and outputs (including images) for up to **30 days**
for abuse and safety monitoring. This applies to all API customers and cannot
be waived via any per-request parameter.

The `store: false` parameter controls only the Stored Completions feature
(distillation/evals). It does not affect the 30-day safety retention window.

OpenAI's Zero Data Retention (ZDR) agreements -- which do eliminate the 30-day
window -- **explicitly exclude image inputs** to the Chat Completions API.
There is currently no contractual path to ZDR coverage for vision calls on
openai.com.

Training opt-out is automatic for all API users since March 1, 2023. No
action is required.

**If strict data residency or full retention control is required for image
inputs**, use Azure OpenAI Service with a regional deployment and the
appropriate Data Processing Agreement. The API surface is nearly identical;
change the endpoint and API key.

---

## Known limitations

- **PDF files are not supported.** Convert PDF pages to images before use.
  `pdftoppm` (Poppler), `Ghostscript`, or `PyMuPDF` can do this.
- **No automatic image resizing.** GPT-5.5 handles images up to 6000px in the
  largest dimension. Very large images are sent as-is; reduce dimensions
  beforehand if needed.
- **Mathematical expressions** may be transcribed incorrectly. Vision models
  can hallucinate arithmetic totals in tables -- verify numeric content
  independently.
- **Handwriting accuracy** varies by legibility. Cursive and non-standard
  letterforms may be misread.
- **Non-Latin scripts** (CJK, Arabic, Devanagari) have lower accuracy than
  Latin-alphabet content. Use `detail: original` (the default with `gpt-5.5`)
  and verify the output.
- **Token limits.** For very dense pages, 4096 output tokens may truncate the
  response. Increase `-MaxTokens` if content appears cut off.
- **No automatic retry.** Any API error aborts the run immediately. Implement
  retry logic in a wrapper script if needed.

---
