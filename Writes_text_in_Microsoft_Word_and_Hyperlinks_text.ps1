# Check if the Windows Capability is installed
$capability = Get-WindowsCapability -Online | Where-Object { $_.Name -eq 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' }

if ($null -ne  $capability) {
    Write-Host "Windows Capability is installed."
} else {
    Write-Host "Windows Capability is not installed."
    Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
}

#Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online

#param([string]$Param1)

#Write-Host "Received parameter value: $Param1"
$user=  Get-Aduser -identity <#username#> -properties Displayname,Description



$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $true

# Create a new document
$document = $wordApp.Documents.Add()

# Set the font of the document to Calibri
$document.Content.Font.Name = "Calibri"

# Add the text with desired formatting
$text = @"
$($user.Displayname)
$($user.description)

inseert text to be written
"@

$document.Content.Text = $text

# Find the position of the word "Facebook"
$range = $document.Content
$matchText = "Facebook"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)

    # Create a hyperlink for the found word
    $hyperlink = $document.Hyperlinks.Add($foundRange, "https://www.facebook.com/", $matchText)
    $hyperlink.Range.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlue
    $hyperlink.Range.Font.Underline = $true
}

# Find the position of the word "Pinterest"
$range = $document.Content
$matchText = "Pinterest"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)

    # Create a hyperlink for the found word
    $hyperlink = $document.Hyperlinks.Add($foundRange, "https://www.pinterest.com/", $matchText)
    $hyperlink.Range.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlue
    $hyperlink.Range.Font.Underline = $true
}

# Find the position of the word "Instagram"
$range = $document.Content
$matchText = "Instagram"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)

    # Create a hyperlink for the found word
    $hyperlink = $document.Hyperlinks.Add($foundRange, "https://www.instagram.com/", $matchText)
    $hyperlink.Range.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlue
    $hyperlink.Range.Font.Underline = $true
}

$range = $document.Content
$matchText = "Insert text to highlight here"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)
    $foundRange.Find.MatchWholeWord = $true
    $foundRange.Bold=$true
  
}

# Find the position of the word www.creativecoop.com"
$range = $document.Content
$matchText = "Insert text to highlight here"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)

    # Create a hyperlink for the found word
    $hyperlink = $document.Hyperlinks.Add($foundRange, "http://www.google.com", $matchText)
    $hyperlink.Range.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlue
    $hyperlink.Range.Font.Underline = $true
}


$range = $document.Content
$matchText = "$($user.DisplayName)"
$found = $range.Find.Execute($matchText)

if ($found) {
    # Get the range of the found word
    $foundRange = $document.Range($range.Start, $range.End)
    $foundRange.Find.MatchWholeWord = $true
    $foundRange.Bold=$true
  }
  
# Save the document
$document.SaveAs("C:\Users\Public\Desktop\your_signature.docx")

# Close the document and Word application
$document.Close()
$wordApp.Quit()
