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
    Gemini and Claude via Azure with Zero Data Retention (ZDR).
    Note: GPT-5.5 support is not yet available on OpenRouter.

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
    Optional. A single model ID to use via OpenRouter. Must support vision
    (image) inputs. Overrides -Models when specified.

.PARAMETER Models
    Optional. An array of model IDs for OpenRouter to try in order. The first
    available model is used; others serve as fallbacks. Defaults to Gemini 3.1 Pro
    Preview, Claude Sonnet 4.6, and Gemini 2.5 Pro. Ignored when -Model
    is used instead.

.PARAMETER Cheapest
    Optional. Route to the least expensive model from -Models instead of using
    the first available. Ignored when -Model is used instead of -Models.

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
    .\Get-OCRTextFromOpenRouter.ps1 -Images scan.png

    Routes to the best available model using the default model list
    (Gemini 3.1 Pro Preview, Claude Sonnet 4.6, Gemini 2.5 Pro).

.EXAMPLE
    .\Get-OCRTextFromOpenRouter.ps1 -Images scan.png -Models "openai/gpt-5.5", "anthropic/claude-sonnet-4"

    Routes to the first available model among the specified options via OpenRouter.
    Note: GPT-5.5 is not yet available on OpenRouter.

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
    [string]$Model,

    [Parameter()]
    [string[]]$Models = @(
        'google/gemini-3.1-pro-preview'
        'anthropic/claude-sonnet-4.6'
        'google/gemini-2.5-pro'
        # 'openai/gpt-5.5' - not yet available on OpenRouter
    ),

    [Parameter()]
    [switch]$Cheapest,

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
        a hashtable with Content, Model, and Cost. Throws on any HTTP error.
        Automatically retries with max_completion_tokens if the first attempt
        fails due to a max_tokens parameter mismatch (e.g. OpenRouter selected
        a GPT-5.x or o-series model).
    #>
    param(
        [string]$ApiKey,
        [string]$Model,
        [string[]]$Models,
        [int]$MaxTokens,
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

    # Sort by price when -Cheapest is set; otherwise OpenRouter tries models in order
    if ($Cheapest -and $Models -and $Models.Count -gt 0) {
        $provider['sort'] = @{
            by        = 'price'
            partition = 'none'
        }
    }

    $body['provider'] = $provider

    # Default to max_tokens; retry with max_completion_tokens on 400 if needed
    $body['max_tokens'] = $MaxTokens
    $body['temperature'] = 0.0

    $jsonBody = $body | ConvertTo-Json -Depth 10

    $response = $null
    $retried = $false

    while ($null -eq $response) {
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

            if ($null -ne $ex.Response) {
                try { $statusCode = [int]$ex.Response.StatusCode } catch {
                    Write-Debug "Could not read HTTP status code: $_"
                }
            }

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

            # Retry with max_completion_tokens on any 400 -- the error detail
            # from the response stream is often empty on PS 5.1, so we cannot
            # reliably check what the API complained about.
            if (-not $retried -and $statusCode -eq 400) {
                Write-Verbose "Retrying with max_completion_tokens (OpenRouter may have selected a GPT-5.x model)"
                $body.Remove('max_tokens')
                $body.Remove('temperature')
                $body['max_completion_tokens'] = $MaxTokens
                $jsonBody = $body | ConvertTo-Json -Depth 10
                $retried = $true
                continue
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
    }

    $cost = 0
    if ($response.usage.cost) { $cost = [double]$response.usage.cost }

    return @{
        Content = $response.choices[0].message.content
        Model   = $response.model
        Cost    = $cost
    }
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
        $result = Invoke-OpenRouterChat `
            -ApiKey $ApiKey `
            -Model $Model `
            -Models $Models `
            -MaxTokens 50 `
            -Messages $classifyMessages
        if ($null -eq $result.Content) { return $false }
        return ($result.Content.Trim().ToUpper() -like 'YES*')
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
$effectiveModels = if (-not [string]::IsNullOrEmpty($Model)) { @($Model) } elseif ($Models -and $Models.Count -gt 0) { $Models } else { @($Model) }
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
$totalCost = 0.0

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

    $result = Invoke-OpenRouterChat `
        -ApiKey $ApiKey `
        -Model $Model `
        -Models $Models `
        -MaxTokens $MaxTokens `
        -Messages $messages.ToArray()

    Write-Verbose "Page $pageIndex | Selected: $($result.Model) | Cost: $('{0:N4}' -f $result.Cost) USD"
    $totalCost += $result.Cost

    $pageMarkdown = $result.Content
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

Write-Verbose "Total cost: $('{0:N4}' -f $totalCost) USD"

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
