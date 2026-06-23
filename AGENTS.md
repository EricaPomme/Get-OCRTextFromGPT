# AGENTS.md

## Project Overview

PowerShell tool that converts document images and chat screenshots to markdown using OpenAI vision models (GPT-5.5 default). Sends images to the Chat Completions API, returns verbatim text transcription with structured markdown for tables, images, charts, and chat transcripts.

## Architecture

```
Get-OCRTextFromGPT.ps1          Entry point - OpenAI API (OPENAI_API_KEY)
Get-OCRTextFromOpenRouterZDR.ps1 Entry point - OpenRouter API (OPENROUTER_API_KEY)
Get-OCRTextFromGPT.Helpers.psm1 Shared module (validation, encoding, model detection, post-processing)
prompts.json                     System prompts, classification prompt, user instruction templates
Add-ContextMenuItems.cmd         Registers Windows Explorer right-click menu entries
Remove-ContextMenuItems.cmd      Removes those entries
```

The two entry-point scripts share their core flow (image validation, EXIF stripping, chat auto-detection, multi-page context, post-processing) but diverge at the API call. The OpenAI script sends `max_completion_tokens` vs `max_tokens` based on `Test-IsGpt5Model`, and the OpenRouter script delegates model selection to OpenRouter via either a single `-Model` or an ordered fallback list `-Models`. The OpenRouter script additionally supports `-Cheapest` (price-sorted routing), automatic 400 retry with `max_completion_tokens` for GPT-5.x / o-series selections, per-page cost logging via the response `usage.cost`, and a total cost summary on completion. Both scripts accept `-Speaker` (chat mode only) to replace the "You" sender label. Shared logic lives in the `.psm1` module.

## Key Gotchas

- **Model family detection controls API parameters**: `Test-IsGpt5Model` in the helpers module checks if the model name matches `gpt-5*`, `o1*`, `o3*`, `o4*`. GPT-5.x/o-series models use `max_completion_tokens` and omit `temperature`. GPT-4o/4.1 models use `max_tokens` with `temperature: 0.0`. Getting this wrong causes API errors.

- **GPT-5.5 gets `detail: original`** (6000px, 10000 patches); all other models get `detail: high`. This is handled by `Get-ImageDetail`.

- **stdout is suppressed when `-ToClipboard` is used** to prevent external tools (like Greenshot) from treating markdown URLs as navigation destinations.

- **Images are re-encoded to PNG before transmission** via `System.Drawing.Bitmap.Save()` to strip EXIF metadata. This is why `.NET Framework 4.x` is required (available on all Windows with PS 5.1).

- **Multi-page context carries forward**: Each page's assistant response is appended to the messages array as an `assistant` message before the next page is processed. This is how cross-page tables/lists work.

- **Chat auto-detection** makes a separate lightweight API call on the first image with the `ChatClassification` prompt (returns YES/NO). Errors silently fall back to document mode.

- **OpenRouter 400 retry**: `Invoke-OpenRouterChat` in the OpenRouter script first sends `max_tokens`; if the response is HTTP 400 it removes that and `temperature`, re-sends with `max_completion_tokens`, and retries once. This handles the case where OpenRouter routed the call to a GPT-5.x or o-series model that rejects `max_tokens`. PS 5.1 often cannot read the 400 body, so retry triggers on any 400 regardless of error detail.

- **OpenRouter cost tracking**: Each `Invoke-OpenRouterChat` return is a hashtable with `Content`, `Model` (the model OpenRouter actually used), and `Cost` (from `response.usage.cost`). The main loop accumulates `Cost` into `$totalCost` and prints per-page and total cost to the verbose stream.

- **OpenRouter ZDR**: The OpenRouter request body sets `provider.zdr = $true`. Effective ZDR coverage depends on whether the routed provider participates in OpenRouter's ZDR program. The OpenRouter script also sets `HTTP-Referer` and `X-Title` headers (required by OpenRouter); it does not set any of the OpenAI privacy headers.

- **Relative timestamp resolution** in chat mode uses the current system date/time. The prompt template in `prompts.json` has positional format placeholders `{0}`-`{4}` filled at runtime.

- **Context menu .cmd scripts use `%~dp0`** to resolve the PowerShell script path relative to the .cmd file's location. The .cmd file must live in the same directory as the .ps1 file.

- **`#Requires -Version 5.1`** on all .ps1/.psm1 files. Targets Windows PowerShell 5.1; also works on PowerShell 7+.

## Prompt Engineering

All prompts live in `prompts.json` and are loaded once via `Get-Prompts` (cached in `$script:PromptsCache`). Key structure:

| Key | Purpose |
|---|---|
| `SystemPrompt` | Document mode system prompt (verbatim transcription rules, markdown formatting, ASCII normalization) |
| `ChatSystemPromptTemplate` | Chat mode template with `{0}`-`{4}` date/time placeholders |
| `ChatClassification` | YES/NO classification prompt for auto-detecting chat screenshots |
| `UserInstructions` | First-page and continuation instructions for both doc and chat modes |
| `SpeakerOverrideTemplate` | Appended to system prompt when `-Speaker` is used; replaces "You" labels |

When modifying prompts: keep the positional format placeholders intact for `ChatSystemPromptTemplate` and `SpeakerOverrideTemplate`.

## Output Post-Processing

Applied to every page response in sequence:
1. `ConvertFrom-CodeFence` - strips leading/trailing code fences (models sometimes wrap entire response)
2. `ConvertTo-AsciiPunctuation` - replaces Unicode smart quotes, em/en dashes, ellipses, non-breaking spaces with ASCII equivalents

## Environment Variables

| Variable | Used by |
|---|---|
| `OPENAI_API_KEY` | `Get-OCRTextFromGPT.ps1` |
| `OPENROUTER_API_KEY` | `Get-OCRTextFromOpenRouterZDR.ps1` |

## File Encoding

All output files are written as **UTF-8 without BOM** using `[System.IO.File]::WriteAllText()` with an explicit `UTF8Encoding($false)`. Do not use PowerShell's `-Encoding utf8` on PS 5.1 (writes UTF-8 with BOM).

## Style

- PowerShell with `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'`
- `[CmdletBinding()]` param blocks on all entry-point scripts
- Verbose output via `Write-Verbose` (does not pollute stdout)
- Progress via `Write-Progress` with parent/child ID hierarchy
- No module manifest (`.psd1`) -- module loaded via `Import-Module` with `-Force`
- No tests, no build system, no CI/CD configured
