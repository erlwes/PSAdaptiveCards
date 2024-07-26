<#
.SYNOPSIS
    Module for creating and managing Adaptive Cards for Microsoft Teams.

.DESCRIPTION
    This module provides functions to generate JSON payloads for Adaptive Cards, Fact Sets, and Tables and more.
    These payloads can be used to post messages to Microsoft Teams channels via Workflow webhooks.

.AUTHOR
    EW

.COPYRIGHT
    No

.LICENSE
    MIT

.VERSION
    0.0.3

.NOTES
    - Requires PowerShell 5.1 or later.
    - For more information, see the Adaptive Cards documentation at https://adaptivecards.io.

.EXAMPLE
    $cardContent = @(
        New-TextBlock -Text "Hello, Teams!"
    )
    New-AdaptiveCard -BodyContent $cardContent | ConvertTo-Json -Depth 20
    
#>

function New-AdaptiveCard {
    param (
        [Parameter(Mandatory = $true)]
        [array]$BodyContent
    )

    $adaptiveCard = [pscustomobject]@{
        type = 'AdaptiveCard'
        body = @()
        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
        version = '1.4'        
    }

    foreach ($item in $BodyContent) {
        $adaptiveCard.body += $item
    }    
    return $adaptiveCard
}

function New-TextBlock {
    param (
        [Parameter(Mandatory = $false)] 
        [bool]$isSubtle,

        [Parameter(Mandatory = $false)]
        [bool]$separator,
        
        [Parameter(Mandatory = $false)]
        [int]$maxLines,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'small', 'medium', 'large', 'extraLarge', IgnoreCase = $false)]
        [string]$size,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'lighter', 'bolder', IgnoreCase = $false)]
        [string]$weight,

        [Parameter(Mandatory = $false)]
        [bool]$wrap,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$color,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'monospace', IgnoreCase = $false)]
        [string]$fontType,

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'none', 'small', 'medium', 'large', 'extraLarge', 'padding', IgnoreCase = $false)]
        [string]$spacing,

        [Parameter(Mandatory = $true)]
        [string]$text
    )
    begin {
        $textBlock = [pscustomobject]@{
            type                = 'TextBlock'
            text                = $text
        }
    }
    process {
        if ($isSubtle)              { $textBlock | Add-Member -NotePropertyName 'isSubtle' -NotePropertyValue $isSubtle }
        if ($separator)             { $textBlock | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
        if ($maxLines)              { $textBlock | Add-Member -NotePropertyName 'maxLines' -NotePropertyValue $maxLines }
        if ($size)                  { $textBlock | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
        if ($weight)                { $textBlock | Add-Member -NotePropertyName 'weight' -NotePropertyValue $weight }
        if ($wrap)                  { $textBlock | Add-Member -NotePropertyName 'wrap' -NotePropertyValue $wrap }
        if ($color)                 { $textBlock | Add-Member -NotePropertyName 'color' -NotePropertyValue $color }
        if ($fonttype)              { $textBlock | Add-Member -NotePropertyName 'fonttype' -NotePropertyValue $fonttype }
        if ($horizontalAlignment)   { $textBlock | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }
        if ($spacing)               { $textBlock | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
    }
    end {
        return $textBlock
    }
}

function New-FactSet {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject]$object,

        [Parameter(Mandatory = $true)]
        [string]$titleProperty,

        [Parameter(Mandatory = $true)]
        [string]$valueProperty
    )    
    begin {
        $facts = @()
    }    
    process {
        $value = [string]$object.$valueProperty
        $fact = @{
            title = $object.$titleProperty
            value = $value
        }
        $facts += $fact
    }
    end {
        $factSet = @{
            type  = 'FactSet'
            facts = $facts
        }
        return $factSet
    }    
}

function New-Table {
    param (        
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject]$object,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')]
        [string]$highlightValueMatch,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')][ValidateSet('dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$highlightValueStyle,

        [Parameter(Mandatory = $false)]
        [bool]$firstRowAsHeader = $true,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$headerRowStyle,
        
        [Parameter(Mandatory = $false)]
        [bool]$showGridLines = $true,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$gridStyle = 'default',

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalCellContentAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('top', 'center', 'bottom', IgnoreCase = $false)]
        [string]$verticalCellContentAlignment
    )
    begin {        
        $table = [pscustomobject]@{
            type                = 'Table'
            gridStyle           = $gridStyle
            firstRowAsHeader    = $firstRowAsHeader
            showGridLines       = $showGridLines
            columns             = @()
            rows                = @()
        }

        # Add optional attributes if provided
        if ($horizontalCellContentAlignment) {
            $table.horizontalCellContentAlignment = $horizontalCellContentAlignment
        }
        if ($verticalCellContentAlignment) {
            $table.verticalCellContentAlignment = $verticalCellContentAlignment
        }

        $columns = @()        
        $isHighlighting = $highlightValueMatch -and $highlightValueStyle
    }
    process {
        if ($columns.Count -eq 0) {

            # Get the noteproperties from the first object
            $columns = $Object.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name

            # Add correct number of columns (one per property)
            foreach ($column in $columns) {
                $table.columns += @{
                    width = 'auto'
                }
            }

            # Add the header row
            $headerRow = @{
                type  = 'TableRow'
                cells = @()
            }

            # Add optional attributes if provided
            if ($headerRowStyle) {
                $headerRow.style = $headerRowStyle
            }
        
            # Add the header row
            foreach ($column in $columns) {
                $headerRow.cells += @{
                    type  = 'TableCell'
                    items = @(
                        @{
                            type = 'TextBlock'
                            text = $column # Noteproperty names
                        }
                    )
                }
            }

            $table.rows += $headerRow
        }

        # Process each object and add a row to the table
        $row = @{
            type  = 'TableRow'
            cells = @()
        }

        foreach ($column in $columns) {
            $textValue = [string]$Object.$column
            $textBlock = @{
                type = 'TextBlock'
                text = $textValue
            }

            if ($isHighlighting -and $textValue -match $highlightValueMatch) {
                $textBlock.color = $highlightValueStyle
            }

            $row.cells += @{
                type  = 'TableCell'
                items = @($textBlock)
            }
        }

        $table.rows += $row
    }
    end {        
        return $table
    }
}

function New-Image {
    param (
        [Parameter(Mandatory = $true)]
        [string]$url,

        [Parameter(Mandatory = $false)]
        [string]$altText = 'image', #Yes, this is mandatory in spec, but getting errors from not providing alt-texts is not amusing, let's be honest.

        [Parameter(Mandatory = $false)]
        [string]$backgroundColor,

        [Parameter(Mandatory = $false)][ValidateScript({$_ -ceq 'auto' -or $_ -ceq 'stretch' -or $_ -match "^\d+px$"}, ErrorMessage = "Height must be either lowercase 'auto', 'stretch', or a number followed by 'px'.")]
        [string]$height = "auto",

        [Parameter(Mandatory = $false)][ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$horizontalAlignment,

        [Parameter(Mandatory = $false)][ValidateSet('auto', 'stretch', 'small', 'medium', 'large', IgnoreCase = $false)]
        [string]$size,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'person', IgnoreCase = $false)]
        [string]$style,

        [Parameter(Mandatory = $false)][ValidatePattern("^\d+(px)?$", ErrorMessage = "Width must be an integer, optionally followed by 'px' to specify this unit.")]
        [string]$width,

        [Parameter(Mandatory = $false)][ValidateSet('default', 'none', 'small', 'medium', 'large', 'extraLarge', 'padding', IgnoreCase = $false)]
        [string]$spacing = 'default',

        [Parameter(Mandatory = $false)]
        [bool]$separator
    )
    begin {
        $imageObject = [PSCustomObject]@{
            type = 'Image'
            url  = $url
            altText  = $altText
        }
    }
    process {
        if ($backgroundColor)       { $imageObject | Add-Member -NotePropertyName 'backgroundColor' -NotePropertyValue $vackgroundColor }
        if ($height -ne 'auto')     { $imageObject | Add-Member -NotePropertyName 'height' -NotePropertyValue $height }
        if ($horizontalAlignment)   { $imageObject | Add-Member -NotePropertyName 'horizontalAlignment' -NotePropertyValue $horizontalAlignment }    
        if ($size)                  { $imageObject | Add-Member -NotePropertyName 'size' -NotePropertyValue $size }
        if ($style)                 { $imageObject | Add-Member -NotePropertyName 'style' -NotePropertyValue $style }
        if ($width)                 { $imageObject | Add-Member -NotePropertyName 'width' -NotePropertyValue $width }
        if ($spacing -ne 'default') { $imageObject | Add-Member -NotePropertyName 'spacing' -NotePropertyValue $spacing }
        if ($separator)             { $imageObject | Add-Member -NotePropertyName 'separator' -NotePropertyValue $separator }
    }
    end {
        return $imageObject
    }
}

function Send-JsonToTeamsWebhook {
    param (
        [Parameter(Mandatory = $true)]
        [string]$webhookURI,

        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [pscustomobject]$adaptiveCard,

        [Parameter(Mandatory = $false)]
        [switch]$fullWidth,

        [Parameter(Mandatory = $false)]
        [switch]$onlyConvertToJson
    )

    $attachment = [pscustomobject]@{
        contentType = 'application/vnd.microsoft.card.adaptive'
        contentUrl = $null
        content = $adaptiveCard
    }

    $message = [pscustomobject]@{
        type = 'message'
        attachments = @()
    }
    $message.attachments += $attachment

    if ($fullWidth) {
        $msteamsProperty = @{
            width = 'Full'
        }
        $message.attachments[0].content | Add-Member -MemberType NoteProperty -Name msteams -Value $msteamsProperty
    }

    $json = ($message | ConvertTo-Json -Depth 20) -replace '\\\\', '\' #-replace "\\", '&#92;'
    
    if ($onlyConvertToJson) {
        Write-Output $json
        Break
    }

    $parameters = @{
        "URI"         = $webhookURI
        "Method"      = 'POST'
        "Body"        = $json
        "ContentType" = 'application/json; charset=UTF-8'
        "ErrorAction" = 'Stop'
    }
    try {
        Invoke-RestMethod @parameters
    }
    catch {
        Write-Error "Failed to send request: $($_.Exception.Message)"
    }
}
