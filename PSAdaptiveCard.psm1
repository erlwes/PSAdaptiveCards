<#
.SYNOPSIS
    Module for creating and managing Adaptive Cards for Microsoft Teams.

.DESCRIPTION
    This module provides functions to generate JSON payloads for Adaptive Cards, Fact Sets, and Tables.
    These payloads can be used to post messages to Microsoft Teams channels via Workflow webhooks.

.AUTHOR
    EW

.COPYRIGHT
    No

.LICENSE
    MIT

.VERSION
    1.0.0

.NOTES
    - Requires PowerShell 5.1 or later.
    - For more information, see the Adaptive Cards documentation at https://adaptivecards.io.

.EXAMPLE
    $cardContent = @(
        New-TextBlock -Text "Hello, Teams!"
    )
    $JSONPayload = New-AdaptiveCard -BodyContent $cardContent
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
        version = '1.6'
        msteams = @{
            width = 'Full'
        }
    }

    foreach ($item in $BodyContent) {
        $adaptiveCard.body += $item
    }
    
    $json = ($adaptiveCard | ConvertTo-Json -Depth 10) -replace '\\\\', '\'
    
    return $json
}

function New-TextBlock {
    param (        
        [bool]$IsSubtle = $false,
        [bool]$separator = $false,        
        [int]$MaxLines,
        [ValidateSet('default', 'small', 'medium', 'large', 'extraLarge', IgnoreCase = $false)]
        [string]$Size = 'default',
        [ValidateSet('default', 'lighter', 'bolder', IgnoreCase = $false)]
        [string]$Weight = 'default',
        [bool]$Wrap = $true,
        [ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention', IgnoreCase = $false)]
        [string]$Color = 'default',
        [ValidateSet('default', 'monospace', IgnoreCase = $false)]
        [string]$Fonttype = 'default',
        [ValidateSet('left', 'center', 'right', IgnoreCase = $false)]
        [string]$HorizontalAlignment = 'left',
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $textBlock = [pscustomobject]@{
        type                = 'TextBlock'
        isSubtle            = $IsSubtle
        separator           = $separator
        maxLines            = $MaxLines
        size                = ($size.ToLower())
        weight              = ($Weight.ToLower())
        wrap                = $Wrap
        color               = $Color
        fonttype            = $Fonttype
        horizontalAlignment = $HorizontalAlignment
        text                = $Text
    }
    return $textBlock
}

function New-FactSet {
    param (
        [Parameter(ValueFromPipeline = $true)]
        [psobject]$Object,
        [Parameter(Mandatory = $true)]
        [string]$TitleProperty,
        [Parameter(Mandatory = $true)]
        [string]$ValueProperty
    )
    
    begin {
        $facts = @()
    }
    
    process {
        $value = [string]$Object.$ValueProperty
        $fact = @{
            title = $Object.$TitleProperty
            value = $value
        }
        $facts += $fact
    }
    end {
        $factSet = @{
            type  = "FactSet"
            facts = $facts
        }
        return $factSet
    }    
}

function New-Table {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(ValueFromPipeline = $true, ParameterSetName = 'Default')]
        [Parameter(ValueFromPipeline = $true, ParameterSetName = 'Highlight')]
        [psobject]$Object,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')]
        [string]$HighlightValueMatch,

        [Parameter(Mandatory = $false, ParameterSetName = 'Highlight')]
        [ValidateSet('dark', 'light', 'accent', 'good', 'warning', 'attention')]
        [string]$HighlightValueStyle,

        [Parameter(Mandatory = $false)]
        [bool]$firstRowAsHeader = $true,

        [Parameter(Mandatory = $false)]
        [ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention')]
        [string]$headerRowStyle,
        
        [Parameter(Mandatory = $false)]
        [bool]$showGridLines = $true,

        [Parameter(Mandatory = $false)]
        [ValidateSet('default', 'dark', 'light', 'accent', 'good', 'warning', 'attention')]
        [string]$gridStyle = 'default',

        [Parameter(Mandatory = $false)]
        [ValidateSet('left', 'center', 'right')]
        [string]$horizontalCellContentAlignment,

        [Parameter(Mandatory = $false)]
        [ValidateSet('top', 'center', 'bottom')]
        [string]$verticalCellContentAlignment
    )

    begin {        
        $table = [pscustomobject]@{
            type                          = 'Table'
            gridStyle                     = $gridStyle
            firstRowAsHeader              = $firstRowAsHeader
            showGridLines                 = $showGridLines
            columns                       = @()
            rows                          = @()
        }

        # Add optional attributes if provided
        if ($horizontalCellContentAlignment) {
            $table.horizontalCellContentAlignment = $horizontalCellContentAlignment
        }
        if ($verticalCellContentAlignment) {
            $table.verticalCellContentAlignment = $verticalCellContentAlignment
        }

        $columns = @()        
        $isHighlighting = $HighlightValueMatch -and $HighlightValueStyle
    }

    process {
        if ($columns.Count -eq 0) {
            # Get the columns from the first object
            $columns = $Object.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name

            # Initialize the columns with equal width
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
        

            foreach ($column in $columns) {
                $headerRow.cells += @{
                    type  = 'TableCell'
                    items = @(
                        @{
                            type = 'TextBlock'
                            text = $column
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

            if ($isHighlighting -and $textValue -eq $HighlightValueMatch) {
                $textBlock.color = $HighlightValueStyle
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

#INSTALL MODULE:
#$modulePath = ($env:PSModulePath -split ';' | Select-Object -First 1)
#New-Item -Type Directory -Path "$modulePath\PSAdaptiveCard"
#New-Item -Type File -Path "$modulePath\PSAdaptiveCard\PSAdaptiveCard.psm1"
#Copy content of this script to '...PSAdaptiveCard.psm1' and save
#Import-Module PSAdaptiveCard
