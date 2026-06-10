#Requires -Version 5.1
<#
.SYNOPSIS
    Converts one or more document or chat images to markdown using the OpenRouter API.

.DESCRIPTION
    Get-OCRTextFromOpenRouter accepts one or more image files (PNG, JPEG, GIF, or WebP),
    sends them to a vision-capable model via the OpenRouter API, and returns
    the extracted content as a single markdown document or chat transcript.

    This script supports the same functionality as Get-OCRTextFromGPT but is designed
    to work with OpenRouter's API, which provides access to various models including
    GPT-5.5 via Azure with Zero Data Retention (ZDR).

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
    uses Zero Data Retention (ZDR) when available. Note that a 30-day safety/abuse
    retention window may still apply on provider servers. See the README for
    full details.

.PARAMETER Images
    One or more paths to image files (PNG, JPEG, GIF, or WebP). When multiple
    images are provided, they are treated as sequential pages of one document.

.PARAMETER OutputPath
    Optional. If specified, the markdown output is also written to this file in
    UTF-8 (no BOM) in addition to being written to stdout.

.PARAMETER ApiKey
    Optional. OpenRouter API key. If not provided, the OPENROUTER_API_KEY environment
    variable is used.

.PARAMETER Model
    Optional. The model to use via OpenRouter. Defaults to 'openai/gpt-5.5:azure-zdr'.
    Must support vision (image) inputs. Ignored when -Models is specified.

.PARAMETER Models
    Optional. An array of model IDs for OpenRouter to route between. When specified,
    OpenRouter selects the best available model based on -SortBy. Overrides -Model.

.PARAMETER SortBy
    Optional. Routing preference when using -Models. Valid values: 'price', 'latency',
    'throughput'. Defaults to 'price' (least expensive). Ignored when -Model is used
    instead of -Models.

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
    .\Get-OCRTextFromOpenRouter.ps1 -Images teams-screenshot.png -ChatMode

    Forces chat transcript mode for a Teams screenshot, skipping auto-detection.

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 teams-p1.png, teams-p2.png -OutputPath transcript.md

    Auto-detects two sequential chat screenshots as a conversation and writes the
    combined transcript to transcript.md.

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 scan.png

    Converts a single image to markdown and writes it to stdout.

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 page1.jpg, page2.jpg, page3.jpg -OutputPath report.md

    Converts three pages of a scanned document and writes the combined result to
    report.md as well as stdout.

.EXAMPLE
    $pages = Get-ChildItem *.png | Sort-Object Name |
             Select-Object -ExpandProperty FullName
    .\Get-OCRTextFromOpenRouter.ps1 -Images $pages -OutputPath combined.md

    Converts all PNG files in the current folder, sorted by name, as pages of
    one document.

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 teams-p1.png, teams-p2.png -ChatMode -Speaker "John Doe"

    Forces chat transcript mode and attributes messages from the local user to
    "John Doe" instead of the default "You" label.

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 -Images scan.png -Models "openai/gpt-5.5", "openai/gpt-4o", "anthropic/claude-sonnet-4" -SortBy price

    Routes to the least expensive vision model among the specified options via OpenRouter.

.NOTES
    Requires an OpenRouter API key with access to a vision-capable model.
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
    [string]$Model = 'openai/gpt-5.5:azure-zdr',

    [Parameter()]
    [string[]]$Models,

    [Parameter()]
    [ValidateSet('price', 'latency', 'throughput')]
    [string]$SortBy = 'price',

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
# Load shared module and prompts
# ---------------------------------------------------------------------------

Import-Module "$PSScriptRoot\Get-OCRTextFromGPT.Helpers.psm1" -Force

$Prompts = Get-Prompts

# ---------------------------------------------------------------------------
# Functions
# ---------------------------------------------------------------------------

function Invoke-OpenRouterChat {
    <#
    .SYNOPSIS
        Sends a messages array to the OpenRouter Chat Completions API and returns
        the assistant's response text. Throws on any HTTP error.
    #>
    param(
        [string]$ApiKey,
        [string]$Model,
        [string[]]$Models,
        [string]$SortBy,
        [int]$MaxTokens,
        [bool]$IsGpt5,
        [array]$Messages
    )

    $headers = @{
        'Authorization'               = "Bearer $ApiKey"
        'HTTP-Referer'                = 'https://github.com/erica/Get-OCRTextFromGPT'
        'X-Title'                     = 'Get-OCRTextFromOpenRouter'
    }

    $body = [ordered]@{
        messages = $Messages
    }

    # Use models array for routing, or single model
    if ($Models -and $Models.Count -gt 0) {
        $body['models'] = $Models
    }
    else {
        $body['model'] = $Model
    }

    # Add ZDR and routing preferences
    $provider = @{
        zdr = $true
    }

    # Add sort preference when using models array
    if ($Models -and $Models.Count -gt 0 -and $SortBy) {
        $provider['sort'] = @{
            by        = $SortBy
            partition = 'none'
        }
    }

    $body['provider'] = $provider

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
            -Uri 'https://openrouter.ai/api/v1/chat/completions' `
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

        # Attempt to read the response body for OpenRouter's error message
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
            throw "OpenRouter API request failed (HTTP $statusCode): $errorDetail"
        }
        elseif ($statusCode) {
            throw "OpenRouter API request failed (HTTP $statusCode)"
        }
        else {
            throw "OpenRouter API request failed: $_"
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
        [string[]]$Models,
        [string]$SortBy,
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
                    text = $Prompts.ChatClassification
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
        $answer = Invoke-OpenRouterChat `
            -ApiKey $ApiKey `
            -Model $Model `
            -Models $Models `
            -SortBy $SortBy `
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
    $ApiKey = $env:OPENROUTER_API_KEY
}
if ([string]::IsNullOrEmpty($ApiKey)) {
    throw 'No API key provided. Set the OPENROUTER_API_KEY environment variable or use the -ApiKey parameter.'
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
$effectiveModels = if ($Models -and $Models.Count -gt 0) { $Models } else { @($Model) }
$isGpt5 = Test-IsGpt5Model -ModelNames $effectiveModels
$detail = Get-ImageDetail -ModelNames $effectiveModels
Write-Verbose "Model(s): $($effectiveModels -join ', ') | GPT-5 parameter set: $isGpt5 | Image detail: $detail"

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
        -Models $Models `
        -SortBy $SortBy `
        -IsGpt5 $isGpt5 `
        -Base64Image $firstImageB64 `
        -Detail $detail
    Write-Verbose "Chat screenshot auto-detected: $useChatMode"
}

if ($useChatMode) {
    $today = Get-Date
    $activeSystemPrompt = $Prompts.ChatSystemPromptTemplate -f `
        $today.ToString('yyyy-MM-dd'), `
        $today.DayOfWeek.ToString(), `
        $today.AddDays(-1).ToString('yyyy-MM-dd'), `
        $today.Year.ToString(), `
        $today.ToString('h:mm tt').ToLower()
    if (-not [string]::IsNullOrEmpty($Speaker)) {
        $activeSystemPrompt += $Prompts.SpeakerOverrideTemplate -f $Speaker
    }
    $progressActivity = 'Converting chat to transcript'
}
else {
    $activeSystemPrompt = $Prompts.SystemPrompt
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
            $instruction = $Prompts.UserInstructions.ChatFirst
        }
        else {
            $instruction = $Prompts.UserInstructions.ChatContinuation
        }
    }
    elseif ($pageIndex -eq 1) {
        $instruction = $Prompts.UserInstructions.DocFirst
    }
    else {
        $instruction = $Prompts.UserInstructions.DocContinuation
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
                    detail = $Detail
                }
            }
        )
    }
    $messages.Add($userMessage)

    Write-Verbose "Processing image $pageIndex of $($totalImages): $imagePath"

    Write-Progress -Id 1 -ParentId 0 -Activity $fileName `
        -Status 'Calling OpenRouter API (this may take a moment)...' -PercentComplete 66

    $pageMarkdown = Invoke-OpenRouterChat `
        -ApiKey $ApiKey `
        -Model $Model `
        -Models $Models `
        -SortBy $SortBy `
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
    Set-ClipboardText -Text $fullMarkdown
    Write-Verbose 'Output copied to clipboard.'
}
