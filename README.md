# PSAdaptiveCards


## Description:
The PSAdaptiveCard PowerShell module designed to create JSON payloads for Adaptive Cards. Adaptive Cards are a way to present content in applications such as Microsoft Teams. This module includes functions to construct various elements of an Adaptive Card, which can be used to post messages to Teams channels using Workflow webhooks.

TLDR; -> [Examples](https://github.com/erlwes/PSAdaptiveCards/tree/main?tab=readme-ov-file#examples-combined)

## Functions:

### New-AdaptiveCard
Description: Creates the structure for an Adaptive Card.

Parameters:

>  _BodyContent_: An array of card elements (e.g., text blocks, fact sets) to be included in the body of the card.

Usage:
```
$cardContent = @(
    New-TextBlock -Text "Hello, Teams!" -weight 'bolder' -size large
    New-TextBlock -Text "Lorem Ipsum Dolar ... "
)
$adaptiveCard = New-AdaptiveCard -BodyContent $cardContent
```

### New-TextBlock
Description: Creates a text block element for an Adaptive Card.

Parameters:
> _IsSubtle_: Boolean indicating if the text should be subtle

> _separator_: Boolean indicating if there should be a separator line

> _MaxLines_: Maximum number of lines to display

> _Size_: Size of the text (default, small, medium, large, extraLarge).

> _Weight_: Weight of the text (default, lighter, bolder).

> _Wrap_: Boolean indicating if the text should wrap.

> _Color_: Color of the text (default, dark, light, accent, good, warning, attention).

> _Fonttype_: Font type of the text (default, monospace).

> _HorizontalAlignment_: Horizontal alignment of the text (left, center, right).

> _Text_ (mandatory): The text content.

Usage:
```
$textBlock = New-TextBlock -Text "This is a text block" -Size "large" -Weight "bolder"
```

### New-FactSet
Description: Creates a fact set element for an Adaptive Card, typically used to present key-value pairs.

Parameters:

> _Object_: Input object from the pipeline

> _TitleProperty_ (mandatory): The property of the object to use as the title

> _ValueProperty_ (mandatory): The property of the object to use as the value

Usage:
```
Get-Service | Select-Object -First 2 | New-FactSet -TitleProperty Name -ValueProperty Status
```

### New-Table
Description: Creates a table element for an Adaptive Card, dynamically defining columns based on object properties and supports highlighting specific cell values.

Parameters:

> _Object_: Input object from the pipeline
 
> _HighlightValueMatch_: Text to match for highlighting
 
> _HighlightValueStyle_: Style to apply for the highlight (dark, light, accent, good, warning, attention)
 
> _firstRowAsHeader_: Boolean to specify if the first row is a header
 
> _headerRowStyle_: Style for the header row
 
> _showGridLines_: Boolean to show grid lines
 
> _gridStyle_: Style for the grid
 
> _horizontalCellContentAlignment_: Horizontal alignment for cell content (left, center, right)

> _verticalCellContentAlignment_: Vertical alignment for cell content (top, center, bottom)

Usage:
```
Get-Service | Select-Object -First 10 | New-Table -HighlightValueMatch "Stopped" -HighlightValueStyle "attention" -firstRowAsHeader $true -showGridLines $false -gridStyle "accent" -horizontalCellContentAlignment "center" -verticalCellContentAlignment "top"
```

## Examples, combined

### Example 1 - Header and a table of services
```
$Services = Get-Service | Select-Object Name, DisplayName, Status -First 5
$cardContent = @(
    New-TextBlock -Size extraLarge -Weight bolder -Text 'Services'
    $Services | New-Table -HighlightValueMatch 'Stopped' -HighlightValueStyle 'attention' -headerRowStyle 'accent' -gridStyle 'accent'
)
New-AdaptiveCard -BodyContent $cardContent | ConvertTo-Json -Depth 20
```
![image](https://github.com/user-attachments/assets/974bc543-54f9-4cee-b840-4f0ff5265e3f)

* Tabled is created from any PowerShell-object first by adding all noteproperties as headers, then adding all objects as additional rows
* Highlighting of matching values of text in textblocks inside of table cells is supported with a parameter set, as illustrated above

### Example 2 - Same as above, but make JSON-payload and send to Teams
```
$Services = Get-Service | Select-Object Name, DisplayName, Status -First 5

$cardContent = @(
    New-TextBlock -Size extraLarge -Weight bolder -Text 'Services'
    $Services | New-Table -HighlightValueMatch 'Stopped' -HighlightValueStyle 'attention' -headerRowStyle 'accent' -gridStyle 'accent'
)

$WebhookURI = 'https://prod-140.westeurope.logic.azure.com:443/workflows/[REDACTED]/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=[REDACTED]'
New-AdaptiveCard -BodyContent $cardContent | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

### Example 3 - Header and a "Fact set"
```
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Employees'
$exampleObjects = @(
    [pscustomobject]@{ Name = 'Jon Doe'; Type = 'Male'; Description = 'Works at Contoso' },
    [pscustomobject]@{ Name = 'Jane Doe'; Type = 'Female'; Description = 'Works at Fabrikam' }
)
$Factset = $exampleObjects | New-FactSet -TitleProperty Name -ValueProperty Description

New-AdaptiveCard -BodyContent $Header, $Factset | ConvertTo-Json -Depth 20
```
![image](https://github.com/user-attachments/assets/3597efea-246f-4bd4-820b-5dd1c10d34b3)


### Example 4 - Header, sub-headers and lists
```
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Good or bad'
$TextBlock1 = New-TextBlock -Size large -Weight bolder -Text 'List 1' -Color attention -separator $true
$TextBlock2 = New-TextBlock -Text '- Item :(\r- Item\r- Item' -Color attention
$TextBlock3 = New-TextBlock -Size large -Weight bolder -Text 'List 2' -Color good -separator $true
$TextBlock4 = New-TextBlock -Text '1. Item :)\r2. Item\r3. Item' -Color good 

New-AdaptiveCard -BodyContent $Header, $TextBlock1, $TextBlock2, $TextBlock3, $TextBlock4 | ConvertTo-Json -Depth 20
```
![image](https://github.com/user-attachments/assets/7dd8cf6c-d1f0-4113-bfa6-a6d35d7e48fd)
