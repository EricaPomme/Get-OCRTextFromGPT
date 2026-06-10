#Requires -Version 5.1
<#
.SYNOPSIS
    Shared helper functions for the Get-OCRTextFromGPT scripts.
.DESCRIPTION
    Provides image validation, base64 encoding, model detection, text
    post-processing, clipboard handling, and prompt loading used by both
    the OpenAI and OpenRouter entry-point scripts.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SupportedExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.webp')
$script:PromptsCache = $null

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
    if ($script:SupportedExtensions -notcontains $ext) {
        throw "Unsupported image format '$ext': $Path -- supported formats: $($script:SupportedExtensions -join ', ')"
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
        penalty parameters. Accepts a single model name or an array.
    #>
    param([string[]]$ModelNames)

    foreach ($name in $ModelNames) {
        $modelNameLower = $name.ToLower()
        $isGpt5 = ($modelNameLower -like 'gpt-5*') -or
        ($modelNameLower -like 'o1*') -or
        ($modelNameLower -like 'o3*') -or
        ($modelNameLower -like 'o4*') -or
        ($modelNameLower -like '*/gpt-5*') -or
        ($modelNameLower -like '*/o1*') -or
        ($modelNameLower -like '*/o3*') -or
        ($modelNameLower -like '*/o4*')
        if ($isGpt5) { return $true }
    }
    return $false
}

function Get-ImageDetail {
    <#
    .SYNOPSIS
        Returns the appropriate image detail level for a given model.
        GPT-5.5 supports "original" (up to 6000px / 10000 patches).
        All other models use "high". Accepts a single model name or an array.
    #>
    param([string[]]$ModelNames)

    foreach ($name in $ModelNames) {
        $modelNameLower = $name.ToLower()
        if ($modelNameLower -like 'gpt-5.5*' -or $modelNameLower -like '*/gpt-5.5*') {
            return 'original'
        }
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

function Set-ClipboardText {
    <#
    .SYNOPSIS
        Copies text to the system clipboard using a dedicated STA runspace.
        Works reliably in both PS 5.1 and PS 7, including when invoked via -File.
    #>
    param([string]$Text)

    $staRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $staRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $staRunspace.Open()
    $staPs = [System.Management.Automation.PowerShell]::Create()
    $staPs.Runspace = $staRunspace
    [void]$staPs.AddScript({
            param($text)
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Clipboard]::SetDataObject($text, $true)
        }).AddArgument($Text)
    [void]$staPs.Invoke()
    $staRunspace.Close()
}

function Get-Prompts {
    <#
    .SYNOPSIS
        Loads prompts.json from the script directory and returns it as a
        hashtable. Caches the result for the lifetime of the PowerShell session.
    #>
    if ($null -ne $script:PromptsCache) {
        return $script:PromptsCache
    }
    $jsonPath = Join-Path $PSScriptRoot 'prompts.json'
    if (-not (Test-Path -LiteralPath $jsonPath -PathType Leaf)) {
        throw "Prompts file not found: $jsonPath"
    }
    $script:PromptsCache = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
    return $script:PromptsCache
}

Export-ModuleMember -Function @(
    'Test-ImageFile',
    'Get-CleanImageBase64',
    'Test-IsGpt5Model',
    'Get-ImageDetail',
    'ConvertFrom-CodeFence',
    'ConvertTo-AsciiPunctuation',
    'Set-ClipboardText',
    'Get-Prompts'
)
