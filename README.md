# PSAdaptiveCards
Generate Adaptive Cards JSON with PowerShell-code

## Example 1 - Header and a table of services
```
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Services'
$Services = Get-Service | Select-Object Name, DisplayName, Status -First 5
$Table = $Services | New-Table -HighlightValueMatch 'Stopped' -HighlightValueStyle 'attention' -headerRowStyle 'accent' -gridStyle 'accent'
New-AdaptiveCard -BodyContent $Header, $Table
```
![image](https://github.com/user-attachments/assets/974bc543-54f9-4cee-b840-4f0ff5265e3f)

* Tabled is created from any PowerShell-object first by adding all noteproperties as headers, then adding all objects as additional rows
* Highlighting of matching values of text in textblocks inside of table cells is supported with a parameter set, as illustrated above

## Example 2 - Header and a "Fact set"
```
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Employees'
$exampleObjects = @(
    [pscustomobject]@{ Name = 'Jon Doe'; Type = 'Male'; Description = 'Works at Contoso' },
    [pscustomobject]@{ Name = 'Jane Doe'; Type = 'Female'; Description = 'Works at Fabrikam' }
)
$Factset = $exampleObjects | New-FactSet -TitleProperty Name -ValueProperty Description

New-AdaptiveCard -BodyContent $Header, $Factset | CLIP
```
![image](https://github.com/user-attachments/assets/3597efea-246f-4bd4-820b-5dd1c10d34b3)


## Example 3 - Header, sub-headers and lists
```
$Header = New-TextBlock -Size extraLarge -Weight bolder -Text 'Good or bad'
$TextBlock1 = New-TextBlock -Size large -Weight bolder -Text 'List 1' -Color attention -separator $true
$TextBlock2 = New-TextBlock -Text '- Item :(\r- Item\r- Item' -Color attention
$TextBlock3 = New-TextBlock -Size large -Weight bolder -Text 'List 2' -Color good -separator $true
$TextBlock4 = New-TextBlock -Text '1. Item :)\r2. Item\r3. Item' -Color good 

New-AdaptiveCard -BodyContent $Header, $TextBlock1, $TextBlock2, $TextBlock3, $TextBlock4
```
![image](https://github.com/user-attachments/assets/7dd8cf6c-d1f0-4113-bfa6-a6d35d7e48fd)
