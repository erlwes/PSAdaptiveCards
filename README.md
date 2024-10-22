# PSAdaptiveCards
The PSAdaptiveCard PowerShell module designed to create JSON-formatted Adaptive Cards. Adaptive Cards are a way to present content in applications such as Microsoft Teams. This module includes functions to construct various elements of an Adaptive Card, and a function for posting these messages to Teams channels using Workflow webhooks.

To get started with worksflows and generate a webhook URL, look [here](https://support.microsoft.com/en-us/office/post-a-workflow-when-a-webhook-request-is-received-in-microsoft-teams-8ae491c7-0394-4861-ba59-055e33f75498#:~:text=An%20Incoming%20webhook%20lets%20external,a%20webhook%20request%20is%20received.&text=next%20to%20the%20channel%20or,for%2C%20and%20then%20select%20Workflows.).


## Straight to the point - an example
```PowerShell
#Use functions to add a header-text and a table to a card.
$cardContent = @(
    New-TextBlock -Size extraLarge -Weight bolder -Text 'Services'
    Get-Service | Select-Object Name, DisplayName, Status -First 5 | New-Table -HighlightValueMatch 'Stopped' -HighlightValueStyle 'attention' -headerRowStyle 'accent' -gridStyle 'accent'
)

#We must provide a webhook-URL. This is a partial example:
$WebhookURI = 'https://prod-140.westeurope.logic.azure.com:443/workflows/[REDACTED]/triggers/manual...'

#Put the elements inside an Adaptive Card, then wrap the card as an attachment to a message, convert it to JSON, and POST it to webhook URL:
New-AdaptiveCard -BodyContent $cardContent | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```
This results in the following message to your Teams-channel:
![image](https://github.com/user-attachments/assets/8ceb598e-2621-4523-bb1c-f674de02a2dc)


## Install
`Install-Module -Name PSAdaptiveCard`


## Functions:

### 🟢New-AdaptiveCard
Creates the structure for an Adaptive Card.

Parameter | Description
--- | ---
BodyContent (mandatory) | An array of card elements (e.g., text blocks, fact sets) to be included in the body of the card `(psobject)`

Usage:
```PowerShell
$cardContent = @(
    New-TextBlock -Text "Hello, Teams!" -weight 'bolder' -size large
    New-TextBlock -Text "Lorem Ipsum Dolar ... "
)
New-AdaptiveCard -BodyContent $cardContent
```

### 🟢Send-JsonToTeamsWebhook
Convert AdaptiveCards to JSON-payload for MS Teams Workflow-based webhook and post it.

Parameter | Description
--- | ---
webhookURI (mandatory) | Specifies the Microsoft Teams webhook URL to which the adaptive card will be sent `(string)`
adaptiveCard (mandatory) | Accept pipeline input, specifies the adaptive card content as a PowerShell custom object `(psobject)`
fullWidth | This switch parameter specifies whether the adaptive card should be displayed at full width in Teams. If this switch is included, the card's width is set to "Full"
onlyConvertToJson | This switch parameter specifies whether the function should only convert the adaptive card to JSON and output it, without sending it to the Teams webhook

Usage:
```PowerShell
$cardContent = @(
    New-TextBlock -Text "Hello, Teams!" -weight 'bolder' -size large
    New-TextBlock -Text "Lorem Ipsum Dolar ... "
)
New-AdaptiveCard -BodyContent $cardContent | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI -fullWidth
```

### 🟢New-TextBlock
Creates a text block element for an Adaptive Card.

Parameter | Description
--- | ---
text (mandatory) | The text content `(string)`
size | Size of the text `(default, small, medium, large, extraLarge)`
weight | Weight of the text `(default, lighter, bolder)`
color | Color of the text `(default, dark, light, accent, good, warning, attention)`
wrap | If true, allow text to wrap. Otherwise, text is clipped `(true, false)`
maxLines | Specifies the maximum number of lines to display. `(int)`
isSubtle | If true, displays text slightly toned down to appear less prominent `(true, false)`
separator | When true, draw a separating line at the top of the element `(true, false)`
spacing | Controls the amount of spacing between this element and the preceding element `(default, none, small, medium, large, extraLarge, padding)`
horizontalAlignment | Horizontal alignment of the text `(left, center, right)`

Usage:
```PowerShell
$textBlock = New-TextBlock -Text "This is a text block" -Size "large" -Weight "bolder"
```

### 🟢New-FactSet
Creates a fact set element for an Adaptive Card, typically used to present key-value pairs.

Parameter | Description
--- | ---
object (mandatory) | Input object from the pipeline `(psobject)`
titleProperty (mandatory) | The property of the object to use as the title `(string)`
valueProperty (mandatory) | The property of the object to use as the value `(string)`

Usage:
```PowerShell
Get-Service | Select-Object -First 2 | New-FactSet -TitleProperty Name -ValueProperty Status
```

### 🟢New-Table
Creates a table element for an Adaptive Card, dynamically defining columns based on object properties and supports highlighting specific cell values.

Parameter | Description
--- | ---
object (mandatory) | Input object from the pipeline `(psobject)`
firstRowAsHeader | Boolean to specify if the first row is a header `(true, false)`
showGridLines | Boolean to show grid lines `(true, false)`
headerRowStyle | Style for the header row `(default, dark, light, accent, good, warning, attention)`
gridStyle | Style for the grid `(default, dark, light, accent, good, warning, attention)`
highlightValueMatch | Text to match for highlighting `(string)`
highlightValueStyle | Style to apply for the highlight `(dark, light, accent, good, warning, attention)`
horizontalCellContentAlignment | Horizontal alignment for cell content `(left, center, right)`
verticalCellContentAlignment | Vertical alignment for cell content `(top, center, bottom)`

Usage:
```PowerShell
Get-Service | Select-Object -First 10 | New-Table -HighlightValueMatch "Stopped" -HighlightValueStyle "attention"
```

### 🟢New-Image
Creates an image-element for an Adaptive Card. Acceptable formats are PNG, JPEG, and GIF.

Parameter | Description
--- | ---
url (mandatory) | The URL to the image `(uri)`
altText | Alternate text describing the image `(string)`
backgroundColor | Applies a background to a transparent image. This property will respect the image style `(string)`
height | The desired height of the image. This overrides the `size` property`
horizontalAlignment | Controls how this element is horizontally positioned within its parent `(left, center, right)`
size | Controls the approximate size of the image `(auto, stretch, small, medium, large)`
style | Controls how this image is displayed `(default, person)`
width | The desired on-screen width of the image, ending in ‘px’. E.g., 50px. This overrides the size property
separator | When true, draw a separating line at the top of the element `(true, false)`
spacing | Controls the amount of spacing between this element and the preceding element `(default, none, small, medium, large, extraLarge, padding)`

Usage:
```PowerShell
New-Image -url "https://adaptivecards.io/content/cats/1.png" -size large -altText 'Cool cat'
```

## Examples

### 🔵Example 1 - Header and a table of services
```PowerShell
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

### 🔵Example 2 - Header and a "Fact set"
```PowerShell
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Employees'
$exampleObjects = @(
    [pscustomobject]@{ Name = 'Jon Doe'; Type = 'Male'; Description = 'Works at Contoso' },
    [pscustomobject]@{ Name = 'Jane Doe'; Type = 'Female'; Description = 'Works at Fabrikam' }
)
$Factset = $exampleObjects | New-FactSet -TitleProperty Name -ValueProperty Description

New-AdaptiveCard -BodyContent $Header, $Factset | ConvertTo-Json -Depth 20
```
![image](https://github.com/user-attachments/assets/3597efea-246f-4bd4-820b-5dd1c10d34b3)


### 🔵Example 3 - Header, sub-headers, lists and emoji's to Teams
```PowerShell
$WebhookURI = '[your webhook uri]'

$Header = New-TextBlock -Size extraLarge -Weight bolder -Text '😇 or 👿'
$TextBlock1 = New-TextBlock -Size large -Weight bolder -Text 'List 1' -Color good -separator $true
$TextBlock2 = New-TextBlock -Text '- Item :)\r- Item\r- Item' -Color good
$TextBlock3 = New-TextBlock -Size large -Weight bolder -Text 'List 2' -Color attention -separator $true
$TextBlock4 = New-TextBlock -Text '1. Item :(\r2. Item\r3. Item' -Color attention 

New-AdaptiveCard -BodyContent $Header, $TextBlock1, $TextBlock2, $TextBlock3, $TextBlock4 | Send-JsonToTeamsWebhook -WebhookURI $WebhookURI
```
![image](https://github.com/user-attachments/assets/f75e6444-a70b-450f-84a4-75a5d9abcde7)
