#Requires -Version 5.1
<#
.SYNOPSIS
    Converts one or more document or chat images to markdown using the OpenAI Vision API.

.DESCRIPTION
    Get-OCRTextFromGPT accepts one or more image files (PNG, JPEG, GIF, or WebP),
    sends them to an OpenAI chat completions model with vision support, and returns
    the extracted content as a single markdown document or chat transcript.

    Multiple images are treated as sequential pages of one document (or sequential
    scrolls of one conversation). Conversation context is carried forward so the
    model can handle cross-page elements such as tables and lists that span page
    boundaries.

    Text is transcribed verbatim. Embedded images and charts are described using
    [Image: ...] and [Chart: ...] notation. Tables are rendered as GitHub-Flavored
    Markdown tables. Blocks of text written entirely in a non-English language are
    followed by an inline machine translation blockquote.

    CHAT MODE: Screenshots of chat or messaging applications (Teams, Slack, Discord,
    etc.) are auto-detected and formatted as transcripts. Each message is rendered
    with a speaker-and-timestamp header. Relative timestamps ("yesterday", "8:50 am")
    are resolved to absolute dates using the current system date. Profile pictures
    and decorative UI elements are omitted; only substantive embedded images are
    described. Use -ChatMode to force transcript mode without auto-detection.

    Privacy: All images are re-encoded as PNG before transmission to strip EXIF
    metadata (GPS coordinates, device identifiers, timestamps). The API request
    sets store=false and includes X-OpenAI-No-Store and X-Stainless-* privacy
    headers. A 30-day safety/abuse retention window still applies on OpenAI's
    servers and cannot be eliminated via API parameters alone. See the README for
    full details.

.PARAMETER Images
    One or more paths to image files (PNG, JPEG, GIF, or WebP). When multiple
    images are provided, they are treated as sequential pages of one document.

.PARAMETER OutputPath
    Optional. If specified, the markdown output is also written to this file in
    UTF-8 (no BOM) in addition to being written to stdout.

.PARAMETER ApiKey
    Optional. OpenAI API key. If not provided, the OPENAI_API_KEY environment
    variable is used.

.PARAMETER Model
    Optional. The OpenAI model to use. Defaults to gpt-5.5. Must support vision
    (image) inputs. GPT-5.x and GPT-4o variants are supported.

.PARAMETER MaxTokens
    Optional. Maximum tokens in the model response per page. Defaults to 4096.

.PARAMETER ChatMode
    Optional. Forces chat transcript mode regardless of auto-detection. Use this
    when processing screenshots of chat or messaging applications (Teams, Slack,
    Discord, etc.) and you want to skip the auto-detection API call. Auto-detection
    runs by default on the first image; -ChatMode bypasses that step.

.PARAMETER Speaker
    Optional. Specifies the name of the local user in chat transcripts. When
    provided, messages that would normally be labeled "You" are instead labeled
    with this name (e.g. "John Doe" instead of "You"). Only applies in chat
    mode. Ignored when processing documents.

.PARAMETER ToClipboard
    Optional. Copies the final markdown output to the system clipboard in addition
    to writing it to stdout.

.EXAMPLE
    .\Get-OCRTextFromGPT.ps1 -Images teams-screenshot.png -ChatMode

    Forces chat transcript mode for a Teams screenshot, skipping auto-detection.

.EXAMPLE
    .\Get-OCRTextFromGPT.ps1 teams-p1.png, teams-p2.png -OutputPath transcript.md

    Auto-detects two sequential chat screenshots as a conversation and writes the
    combined transcript to transcript.md.

.EXAMPLE
    .\Get-OCRTextFromGPT.ps1 scan.png

    Converts a single image to markdown and writes it to stdout.

.EXAMPLE
    .\Get-OCRTextFromGPT.ps1 page1.jpg, page2.jpg, page3.jpg -OutputPath report.md

    Converts three pages of a scanned document and writes the combined result to
    report.md as well as stdout.

.EXAMPLE
    $pages = Get-ChildItem *.png | Sort-Object Name |
             Select-Object -ExpandProperty FullName
    .\Get-OCRTextFromGPT.ps1 -Images $pages -OutputPath combined.md

    Converts all PNG files in the current folder, sorted by name, as pages of
    one document.

.EXAMPLE
    .\Get-OCRTextFromGPT.ps1 teams-p1.png, teams-p2.png -ChatMode -Speaker "John Doe"

    Forces chat transcript mode and attributes messages from the local user to
    "John Doe" instead of the default "You" label.

.NOTES
    Requires an OpenAI API key with access to a vision-capable model.
    Requires System.Drawing, which is available on all Windows systems with
    .NET Framework 4.x (included with PowerShell 5.1).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Images,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$ApiKey,

    [Parameter()]
    [string]$Model = 'gpt-5.5',

    [Parameter()]
    [int]$MaxTokens = 4096,

    [Parameter()]
    [switch]$ChatMode,

    [Parameter()]
    [string]$Speaker,

    [Parameter()]
    [switch]$ToClipboard
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$SupportedExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.webp')

$SystemPrompt = @'
You are a precise document conversion assistant. Your task is to convert document images to well-structured markdown.

REQUIRED RULES:

1. Transcribe all visible substantive text faithfully. Preserve capitalization, punctuation, paragraph breaks, list indentation, and meaningful spacing. Do not preserve incidental spacing caused by text justification, OCR segmentation, scan artifacts, or automatic visual wrapping unless the spacing carries meaning.

2. Normalize English typography according to the ASCII rule below. This means the output is a faithful, machine-friendly transcription, not a typographic facsimile.

3. Use correct markdown syntax for headings (#, ##, ###), lists (*, -, 1.), bold (**text**), italic (*text*), and inline code (`code`) when these structures are present or strongly implied by the document layout.

4. Preserve the logical reading order of the content. For multi-column documents, read each column in its natural order before moving to the next column, unless the layout clearly indicates another reading order.

5. Render clear tables as GitHub-Flavored Markdown tables with alignment separators. Preserve column headers and row order.

6. If a table's structure is unclear, do not guess the columns. Preserve the table-like content in a fenced plain-text block instead. Mark unreadable cells or passages as [...].

7. Render structured form fields (label: value pairs) as a two-column markdown table by default:

   | Field | Value |
   | ----- | ----- |

8. Render code, command-line text, configuration, log output, or visibly monospaced content in code fences. Include a language identifier where determinable.

9. Describe embedded images, illustrations, or photographs using this notation:
   [Image: <description>]
   Be factual and concise. Use dot-separated phrases. Do not speculate.
   Example: [Image: Portrait photograph. Middle-aged person in formal attire. Light background.]

10. Describe charts, graphs, and diagrams using this notation:
    [Chart: <description>]
    Use factual dot-separated phrases. Include chart type, axis labels, legend labels, and visible trend or relationship when legible.
    Example: [Chart: Bar chart. X-axis: years 2019-2023. Y-axis: revenue in millions. Five bars showing an increasing trend.]

11. If text inside an embedded image, chart, or screenshot is too small or unclear to read, say so in the description. Do not invent unreadable text.

12. Mark illegible or obscured individual characters as [?]. Mark longer unreadable passages as [...].

13. Do not include running headers, footers, page numbers, watermarks, crop marks, decorative rules, or repeated boilerplate unless they carry substantive content.

14. Do not add commentary, summaries, explanations, confidence notes, or any text that is not part of the document, except for required [Image: ...], [Chart: ...], illegibility markers, and translation blockquotes.

15. Preserve footnotes, endnotes, marginal notes, and citation markers. Use standard markdown footnote syntax where possible. If note text appears at the bottom of the page, place it after the relevant paragraph or at the end of the page.

16. When a complete paragraph or block of text is written entirely in a language other than English, output that block verbatim, then immediately follow it with a translation blockquote in exactly this format:

    > *[Machine translation from <Language>: <translated text>]*

    Do not add a translation blockquote for single words, short phrases, loan words, proper nouns, technical terms, or foreign-language expressions naturally embedded within otherwise English text. Translate only self-contained blocks that are wholly in another language.

17. This rule applies to English text only: when outputting English, use only plain ASCII characters (code points 0-127). Use straight apostrophes (') not curly/smart apostrophes; straight double quotes (") not typographic quotes; a hyphen (-) or spaced hyphen ( - ) not en/em dashes; three dots (...) not the ellipsis character. Non-English text - Japanese, French, Arabic, accented Latin, Cyrillic, etc. - must be transcribed faithfully in its native characters.

18. Output only the converted markdown content.
'@


# Template for the chat transcript system prompt.
# Placeholders filled at runtime by the -f operator:
# {0} = today's date (yyyy-MM-dd)  {1} = day of week
# {2} = yesterday's date           {3} = current year
# {4} = current time

$ChatSystemPromptTemplate = @'
You are a precise chat transcript formatter. Your task is to convert screenshots of chat or messaging applications into clean, readable markdown transcripts.

Reference date: Today is {0} ({1}). Yesterday was {2}. Current time is {4}.

REQUIRED RULES:

1. Format each message as:
   **Sender Name** (YYYY-MM-DD H:MM am/pm)

   Message content here.

   Leave one blank line between messages.

2. Resolve all timestamps to absolute dates using the reference date above.
   Priority order:
   (a) An explicit date visible in the screenshot takes precedence over everything.
   (b) The chronological ordering of messages, ascending, narrows possibilities.
   (c) Apply the rules below as best-effort inference only when no explicit date is available.

   * A bare time with no date context (e.g. "8:50 am"): compare it to the current time ({4}). If the time is on or before {4}, assign today ({0}). If the time is later than {4}, that hour has not yet occurred today, so assign yesterday ({2}).
   * "Yesterday" resolves to {2}.
   * Named days ("Monday", "Tuesday", etc.) resolve to the most recent occurrence of that weekday on or before today.
   * Month-and-day dates (e.g. "May 26") use the current year ({3}) unless that date would be in the future, in which case use the prior year.
   * Full timestamps already showing a complete date and time should be reformatted into the standard format above without changing the date or time.
   * Do not invent dates outside these rules.

3. Transcribe all message text faithfully. Preserve capitalization, punctuation, paragraph breaks, intentional line breaks, and emoji. Do not preserve automatic visual wrapping caused by screen width.

4. Normalize English typography according to the ASCII rule below. This means the output is a faithful, machine-friendly transcript, not a typographic facsimile.

5. Determine sender names from visible UI labels and message layout.

   * If the screenshot identifies the local user's messages as "You", use **You** exactly.
   * If the local user's name is not visible but the message position or UI styling clearly indicates the local user, use **You**.
   * If a sender name is visible, use the visible name exactly as shown, normalized only according to the ASCII rule if the name is English.
   * If the sender cannot be determined from visible text or layout, use **Unknown**.
   * Never infer a sender name from nearby messages unless the UI clearly groups that message under that sender.

6. Group text into a single message only when it is visually part of the same message bubble/card or clearly under the same sender/timestamp grouping.

   * If a new visible timestamp appears, start a new message.
   * If a new sender appears, start a new message.
   * If a sender/timestamp applies to a block of multiple paragraphs, keep those paragraphs under one message unless a new sender, timestamp, or message boundary appears.

7. If a message is visibly cut off by a "see more", "read more", or similar expansion link, include all visible text and then append on a new line:
   [message continues -- truncated in screenshot]

8. Skip non-content UI chrome: application sidebar, toolbar, window title bar, channel/room name banners, notification badges, read receipts, typing indicators, and timestamp separators that show only a date with no message.

9. Skip decorative and trivial visual elements: profile picture thumbnails, small avatars, status dots, presence icons, app icons, and other small UI decorations. Do not describe these using [Image: ...].

10. Do not describe message reaction icons as images. If the UI shows a message reaction count, transcribe it using the required reaction line.

    * If the reaction emoji is recognizable, output the actual emoji character.
      Example: *Reactions: 👍 x 1*
    * If the emoji is visible but not identifiable, use:
      *Reactions: [unknown emoji] x N*
    * If only a reaction count is visible but the emoji is not visible, use:
      *Reactions: [reaction] x N*

11. Describe only genuinely substantive embedded content, such as a full-size shared photo, document preview card, screenshot, chart, or file preview posted within the conversation.

12. Use this notation for substantive embedded images:
    [Image: <description>]

    Prefer boring, standardized descriptions. Use the broad content type unless the embedded text is legible. Examples:
    [Image: Screenshot of an email thread. Text is too small to read.]
    [Image: Screenshot of a document.]
    [Image: Screenshot of a spreadsheet.]
    [Image: Screenshot of a message details dialog.]
    [Image: Shared photograph.]

13. If the embedded image contains legible substantive text that is part of the conversation, summarize the visible content concisely in the image description. If the text is too small or unclear to read, say so. Do not invent unreadable text.

14. When a complete message is written entirely in a language other than English, output the message verbatim, then immediately follow it with a translation blockquote:

    > *[Machine translation from <Language>: <translated text>]*

    Do not translate single words, names, short expressions, loan words, or foreign-language expressions naturally embedded within otherwise English text.

15. If multiple images are provided, treat them as sequential scrolls of the same conversation in the order given.

    * Do not repeat messages that already appeared in a previous screenshot.
    * When overlapping screenshots show the same message, output it once using the most complete visible version.
    * If one screenshot cuts off a message and a later screenshot shows the full message, use the full later version.

16. This rule applies to English text only: when outputting English, use only plain ASCII characters (code points 0-127). Use straight apostrophes (') not curly/smart apostrophes; straight double quotes (") not typographic quotes; a hyphen (-) or spaced hyphen ( - ) not en/em dashes; three dots (...) not the ellipsis character. Non-English text - Japanese, French, Arabic, accented Latin, Cyrillic, etc. - must be transcribed faithfully in its native characters.

17. Do not add headings, summaries, commentary, confidence notes, or any text not present in the original screenshot, except for required [Image: ...], reaction lines, truncation markers, illegibility markers, and translation blockquotes.

18. Output only the faithful markdown transcript.
'@


# ---------------------------------------------------------------------------
# Functions
# ---------------------------------------------------------------------------

function Test-ImageFile {
    <#
    .SYNOPSIS
        Validates that a file exists and has a supported image extension.
        Throws a descriptive message on failure.
    #>
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Image file not found: $Path"
    }
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    if ($SupportedExtensions -notcontains $ext) {
        throw "Unsupported image format '$ext': $Path -- supported formats: $($SupportedExtensions -join ', ')"
    }
}

function Get-CleanImageBase64 {
    <#
    .SYNOPSIS
        Loads an image via System.Drawing, re-encodes it as PNG (stripping all
        EXIF metadata), and returns the base64-encoded bytes as a string.
    #>
    param([string]$Path)

    $absPath = (Resolve-Path -LiteralPath $Path).Path
    $bitmap = $null
    $ms = $null
    try {
        $bitmap = [System.Drawing.Bitmap]::new($absPath)
        $ms = [System.IO.MemoryStream]::new()
        $bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        return [System.Convert]::ToBase64String($ms.ToArray())
    }
    finally {
        if ($null -ne $ms) { $ms.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
}

function Test-IsGpt5Model {
    <#
    .SYNOPSIS
        Returns $true for models that use the GPT-5.x / o-series parameter set:
        max_completion_tokens instead of max_tokens, and no temperature or
        penalty parameters.
    #>
    param([string]$ModelName)

    return ($ModelName -like 'gpt-5*') -or
    ($ModelName -like 'o1*') -or
    ($ModelName -like 'o3*') -or
    ($ModelName -like 'o4*')
}

function Get-ImageDetail {
    <#
    .SYNOPSIS
        Returns the appropriate image detail level for a given model.
        GPT-5.5 supports "original" (up to 6000px / 10000 patches).
        All other models use "high".
    #>
    param([string]$ModelName)

    if ($ModelName -like 'gpt-5.5*') {
        return 'original'
    }
    return 'high'
}

function ConvertFrom-CodeFence {
    <#
    .SYNOPSIS
        Strips a leading ```[language] fence and a trailing ``` fence from a
        string. Models sometimes wrap their entire response in a code block.
    #>
    param([string]$Text)

    $Text = $Text -replace '^```[a-zA-Z]*\r?\n', ''
    $Text = $Text -replace '\r?\n```\s*$', ''
    return $Text
}

function ConvertTo-AsciiPunctuation {
    <#
    .SYNOPSIS
        Replaces common Unicode typographic characters with plain ASCII
        equivalents. Prevents encoding garbling when the model silently
        upgrades straight quotes, dashes, or ellipses to Unicode forms.
    #>
    param([string]$Text)

    # Smart/curly single quotes -> straight apostrophe
    $Text = $Text.Replace([string][char]0x2018, "'").Replace([string][char]0x2019, "'")
    # Smart/curly double quotes -> straight double quote
    $Text = $Text.Replace([string][char]0x201C, '"').Replace([string][char]0x201D, '"')
    # Em dash -> spaced hyphen
    $Text = $Text.Replace([string][char]0x2014, ' - ')
    # En dash -> hyphen
    $Text = $Text.Replace([string][char]0x2013, '-')
    # Horizontal ellipsis -> three dots
    $Text = $Text.Replace([string][char]0x2026, '...')
    # Non-breaking space -> regular space
    $Text = $Text.Replace([string][char]0x00A0, ' ')
    return $Text
}

function Invoke-OpenAIChat {
    <#
    .SYNOPSIS
        Sends a messages array to the OpenAI Chat Completions API and returns
        the assistant's response text. Throws on any HTTP error.
    #>
    param(
        [string]$ApiKey,
        [string]$Model,
        [int]$MaxTokens,
        [bool]$IsGpt5,
        [array]$Messages
    )

    $headers = @{
        'Authorization'               = "Bearer $ApiKey"
        'X-OpenAI-No-Store'           = 'true'
        'X-Stainless-OS'              = 'private'
        'X-Stainless-Arch'            = 'private'
        'X-Stainless-Runtime'         = 'private'
        'X-Stainless-Runtime-Version' = 'private'
    }

    $body = [ordered]@{
        model    = $Model
        store    = $false
        messages = $Messages
    }

    if ($IsGpt5) {
        $body['max_completion_tokens'] = $MaxTokens
    }
    else {
        $body['max_tokens'] = $MaxTokens
        $body['temperature'] = 0.0
    }

    $jsonBody = $body | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod `
            -Uri 'https://api.openai.com/v1/chat/completions' `
            -Method Post `
            -Headers $headers `
            -Body $jsonBody `
            -ContentType 'application/json'
    }
    catch {
        $ex = $_.Exception
        $statusCode = $null
        $errorDetail = ''

        # Resolve HTTP status code from the response (works on both PS 5.1 and PS 7)
        if ($null -ne $ex.Response) {
            try { $statusCode = [int]$ex.Response.StatusCode } catch {
                Write-Debug "Could not read HTTP status code: $_"
            }
        }

        # Attempt to read the response body for OpenAI's error message
        if ($null -ne $ex.Response) {
            try {
                $stream = $ex.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($stream)
                $errorDetail = $reader.ReadToEnd()
                $reader.Dispose()
                $stream.Dispose()
            }
            catch {
                Write-Debug "Could not read response body: $_"
            }
        }

        if ($statusCode -and $errorDetail) {
            throw "OpenAI API request failed (HTTP $statusCode): $errorDetail"
        }
        elseif ($statusCode) {
            throw "OpenAI API request failed (HTTP $statusCode)"
        }
        else {
            throw "OpenAI API request failed: $_"
        }
    }

    return $response.choices[0].message.content
}

function Test-IsChatScreenshot {
    <#
    .SYNOPSIS
        Makes a lightweight API call to classify whether the provided image is a
        screenshot of a chat or messaging application. Returns $true if so.
        On any error, returns $false (falls back to document mode silently).
    #>
    param(
        [string]$ApiKey,
        [string]$Model,
        [bool]$IsGpt5,
        [string]$Base64Image,
        [string]$Detail
    )

    $classifyMessages = @(
        @{
            role    = 'user'
            content = @(
                @{
                    type = 'text'
                    text = 'Examine this screenshot carefully. Does it show a chat or messaging conversation? Indicators include: a vertical thread of individual short messages, sender names or avatars beside each message, per-message timestamps, message bubbles or bordered text blocks, and a reply input box at the bottom. Applications that match this pattern include Microsoft Teams, Slack, Discord, WhatsApp, iMessage, Google Chat, and similar. Reply with only the single word YES or NO.'
                },
                @{
                    type      = 'image_url'
                    image_url = @{
                        url    = "data:image/png;base64,$Base64Image"
                        detail = $Detail
                    }
                }
            )
        }
    )

    try {
        $answer = Invoke-OpenAIChat `
            -ApiKey $ApiKey `
            -Model $Model `
            -MaxTokens 50 `
            -IsGpt5 $IsGpt5 `
            -Messages $classifyMessages
        return ($answer.Trim().ToUpper() -like 'YES*')
    }
    catch {
        Write-Verbose "Chat auto-detection failed, defaulting to document mode: $_"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

# Resolve API key
if ([string]::IsNullOrEmpty($ApiKey)) {
    $ApiKey = $env:OPENAI_API_KEY
}
if ([string]::IsNullOrEmpty($ApiKey)) {
    throw 'No API key provided. Set the OPENAI_API_KEY environment variable or use the -ApiKey parameter.'
}

# Validate all image files before making any API calls
foreach ($imagePath in $Images) {
    Test-ImageFile -Path $imagePath
}
Write-Verbose "Validated $($Images.Count) image(s)."

# Show progress immediately so the user knows the script has started
Write-Progress -Id 0 -Activity 'Converting document to markdown' `
    -Status 'Initializing...' -PercentComplete 0

# Load System.Drawing for EXIF stripping (available via .NET Framework on all Windows systems)
Add-Type -AssemblyName System.Drawing

# Determine model characteristics once for all pages
$isGpt5 = Test-IsGpt5Model -ModelName $Model
$detail = Get-ImageDetail -ModelName $Model
Write-Verbose "Model: $Model | GPT-5 parameter set: $isGpt5 | Image detail: $detail"

# Determine which system prompt to use; auto-detect chat screenshots unless -ChatMode is set
$useChatMode = $false
$firstImageB64 = $null

if ($ChatMode.IsPresent) {
    $useChatMode = $true
    Write-Progress -Id 0 -Activity 'Converting document to markdown' `
        -Status 'Preparing chat transcript mode...' -PercentComplete 0
}
else {
    Write-Progress -Id 0 -Activity 'Converting document to markdown' `
        -Status 'Detecting image type...' -PercentComplete 0
    $firstImageB64 = Get-CleanImageBase64 -Path $Images[0]
    $useChatMode = Test-IsChatScreenshot `
        -ApiKey $ApiKey `
        -Model $Model `
        -IsGpt5 $isGpt5 `
        -Base64Image $firstImageB64 `
        -Detail $detail
    Write-Verbose "Chat screenshot auto-detected: $useChatMode"
}

if ($useChatMode) {
    $today = Get-Date
    $activeSystemPrompt = $ChatSystemPromptTemplate -f `
        $today.ToString('yyyy-MM-dd'), `
        $today.DayOfWeek.ToString(), `
        $today.AddDays(-1).ToString('yyyy-MM-dd'), `
        $today.Year.ToString(), `
        $today.ToString('h:mm tt').ToLower()
    if (-not [string]::IsNullOrEmpty($Speaker)) {
        $activeSystemPrompt += @"

SPEAKER OVERRIDE: The local user's name is "$Speaker". Replace all occurrences of "You" as a speaker name with "$Speaker" throughout the transcript. For example, "**You** (2025-06-10 3:15 pm)" becomes "**$Speaker** (2025-06-10 3:15 pm)". Do not change "you" when it appears as a regular word inside message content.
"@
    }
    $progressActivity = 'Converting chat to transcript'
}
else {
    $activeSystemPrompt = $SystemPrompt
    $progressActivity = 'Converting document to markdown'
}

# Initialize conversation with the system prompt
$messages = [System.Collections.Generic.List[hashtable]]::new()
$messages.Add(@{
        role    = 'system'
        content = $activeSystemPrompt
    })

# Process each image as the next page of a continuous document or conversation
$pageResults = [System.Collections.Generic.List[string]]::new()
$pageIndex = 0
$totalImages = $Images.Count

foreach ($imagePath in $Images) {
    $pageIndex++
    $fileName = [System.IO.Path]::GetFileName($imagePath)
    $pctComplete = [int](($pageIndex - 1) / $totalImages * 100)

    Write-Progress -Id 0 -Activity $progressActivity `
        -Status "Image $pageIndex of $($totalImages): $fileName" `
        -PercentComplete $pctComplete

    Write-Progress -Id 1 -ParentId 0 -Activity $fileName `
        -Status 'Stripping EXIF metadata...' -PercentComplete 33

    # Use cached base64 for the first image if auto-detection already encoded it
    if ($pageIndex -eq 1 -and $null -ne $firstImageB64) {
        $b64 = $firstImageB64
    }
    else {
        $b64 = Get-CleanImageBase64 -Path $imagePath
    }

    if ($useChatMode) {
        if ($pageIndex -eq 1) {
            $instruction = 'Convert this chat screenshot to a markdown transcript.'
        }
        else {
            $instruction = 'This is the next screenshot in the same conversation. Continue the transcript from where you left off.'
        }
    }
    elseif ($pageIndex -eq 1) {
        $instruction = 'Convert this document image to markdown.'
    }
    else {
        $instruction = 'This is the next page of the same document. Continue the markdown conversion without repeating the document title or any headings already established.'
    }

    $userMessage = @{
        role    = 'user'
        content = @(
            @{
                type = 'text'
                text = $instruction
            },
            @{
                type      = 'image_url'
                image_url = @{
                    url    = "data:image/png;base64,$b64"
                    detail = $detail
                }
            }
        )
    }
    $messages.Add($userMessage)

    Write-Verbose "Processing image $pageIndex of $($totalImages): $imagePath"

    Write-Progress -Id 1 -ParentId 0 -Activity $fileName `
        -Status 'Calling OpenAI API (this may take a moment)...' -PercentComplete 66

    $pageMarkdown = Invoke-OpenAIChat `
        -ApiKey $ApiKey `
        -Model $Model `
        -MaxTokens $MaxTokens `
        -IsGpt5 $isGpt5 `
        -Messages $messages.ToArray()

    $pageMarkdown = ConvertFrom-CodeFence -Text $pageMarkdown
    $pageMarkdown = ConvertTo-AsciiPunctuation -Text $pageMarkdown

    Write-Progress -Id 1 -ParentId 0 -Activity $fileName -Completed

    # Carry the assistant response forward as context for subsequent pages
    $messages.Add(@{
            role    = 'assistant'
            content = $pageMarkdown
        })

    $pageResults.Add($pageMarkdown.TrimEnd())
}

# Combine all pages into one continuous document
$fullMarkdown = $pageResults.ToArray() -join "`n`n"

Write-Progress -Id 0 -Activity $progressActivity -Completed

# Write to stdout unless -ToClipboard is specified.
# When -ToClipboard is used, suppress stdout to prevent external tools (like Greenshot)
# from parsing URLs in the markdown and treating them as navigation destinations.
if (-not $ToClipboard) {
    Write-Output $fullMarkdown
}

# Write to file if requested (UTF-8 without BOM for cross-tool compatibility)
if (-not [string]::IsNullOrEmpty($OutputPath)) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OutputPath, $fullMarkdown, $utf8NoBom)
    Write-Verbose "Output written to: $OutputPath"
}

if ($ToClipboard) {
    # Set-Clipboard hangs when invoked via -File (MTA mode in PS 5.1).
    # clip.exe mangles Unicode (emoji, etc.) due to OEM code-page conversion.
    # Solution: a dedicated STA runspace with Windows Forms.
    # SetDataObject(..., $true) flushes data to global memory (OleFlushClipboard)
    # so the clipboard contents survive after this process exits.
    $staRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $staRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $staRunspace.Open()
    $staPs = [System.Management.Automation.PowerShell]::Create()
    $staPs.Runspace = $staRunspace
    [void]$staPs.AddScript({
            param($text)
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Clipboard]::SetDataObject($text, $true)
        }).AddArgument($fullMarkdown)
    [void]$staPs.Invoke()
    $staRunspace.Close()
    Write-Verbose 'Output copied to clipboard.'
}
