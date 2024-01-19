# Hardcoded path to the HTML Agility Pack DLL
$htmlAgilityPackPath = "PATH TO YOUR DLL"

# Load the HTML Agility Pack assembly
Add-Type -Path $htmlAgilityPackPath

# Function to extract information from HTML using HTML Agility Pack
function Extract-InfoWithHtmlAgilityPack {
    param (
        [string]$htmlContent,
        [string]$label
    )

    # Create an HTML Agility Pack document
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    # Select the element with the specified label
    $element = $doc.DocumentNode.SelectSingleNode("//span[@class='label' and contains(text(), '$label')]/following-sibling::strong[@class='value']")

    # If the element is not found, try to find the span with the specified label
    if (-not $element) {
        $element = $doc.DocumentNode.SelectSingleNode("//span[@class='label' and contains(text(), '$label')]/following-sibling::span[@class='value']")
    }

    # Output the inner text of the selected element
    if ($element) {
        $element.InnerText
    } else {
        "Not Found"
    }
}

# Function to extract badge information from HTML using HTML Agility Pack
function Extract-BadgeInfoWithHtmlAgilityPack {
    param (
        [string]$htmlContent
    )

    # Create an HTML Agility Pack document
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    # Select the badge information within the specified structure
    $badgeElement = $doc.DocumentNode.SelectSingleNode("//p[@class='entry' and .//span[@class='icon']]//span[@class='value']")

    # Output the inner text of the selected badge element
    if ($badgeElement) {
        $badgeElement.InnerText
    } else {
        "Not Found"
    }
}

# Function to extract Enlisted date information from HTML using HTML Agility Pack
function Extract-EnlistedDateWithHtmlAgilityPack {
    param (
        [string]$htmlContent
    )

    # Create an HTML Agility Pack document
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    # Select the Enlisted date information within the specified structure
    $enlistedDateElement = $doc.DocumentNode.SelectSingleNode("//p[@class='entry' and .//span[@class='label' and contains(text(), 'Enlisted')]]/strong[@class='value']")

    # Output the inner text of the selected Enlisted date element
    if ($enlistedDateElement) {
        $enlistedDateElement.InnerText
    } else {
        "Not Found"
    }
}

# Function to extract Fluency information from HTML using HTML Agility Pack
function Extract-FluencyWithHtmlAgilityPack {
    param (
        [string]$htmlContent
    )

    # Create an HTML Agility Pack document
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    # Select the Fluency information within the specified structure
    $fluencyElement = $doc.DocumentNode.SelectSingleNode("//p[@class='entry' and .//span[@class='label' and contains(text(), 'Fluency')]]/strong[@class='value']")

    # Output the inner text of the selected Fluency element after replacing multiple spaces with a single space
    if ($fluencyElement) {
        $fluency = $fluencyElement.InnerText -replace '\s+', ' '
        $fluency.Trim()
    } else {
        "Not Found"
    }
}

# Function to get organization details
function Get-OrganizationDetails {
    param (
        [string]$htmlContent
    )

    # Use HTML Agility Pack to load the HTML content
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    # Extract main organization details
    $mainOrg = $doc.DocumentNode.SelectSingleNode("//div[contains(@class, 'box-content org main')]")
    if ($mainOrg) {
        $visibilityClass = $mainOrg.GetAttributeValue("class", "")
        $isRedacted = $visibilityClass -like "*visibility-R*"

        if ($isRedacted) {
            Write-Host -NoNewline "Main Organization: "; Write-Host -ForegroundColor DarkRed "REDACTED"
        } else {
            $mainOrgNameNode = $mainOrg.SelectSingleNode(".//a[@class='value']")
            $mainOrgName = if ($mainOrgNameNode) { $mainOrgNameNode.InnerText.Trim() } else { "" }

            $mainOrgSIDNode = $mainOrg.SelectSingleNode(".//strong[@class='value']")
            $mainOrgSID = if ($mainOrgSIDNode) { $mainOrgSIDNode.InnerText.Trim() } else { "" }

            Write-Host -NoNewline "Main Organization: "; Write-Host -ForegroundColor DarkCyan "$mainOrgName($mainOrgSID)"
        }
    } else {
        Write-Host -NoNewline "Main Organization: "; Write-Host -ForegroundColor DarkCyan "None"
    }

    # Extract affiliated organization details
    $affiliatedOrgs = $doc.DocumentNode.SelectNodes("//div[contains(@class, 'box-content org affiliation')]")
    $affiliatedOrgDetails = @()

    if ($affiliatedOrgs) {
        foreach ($affiliatedOrg in $affiliatedOrgs) {
            $visibilityClass = $affiliatedOrg.GetAttributeValue("class", "")
            $isRestricted = $visibilityClass -like "*visibility-R*"

            if ($isRestricted) {
                $affiliatedOrgDetails += "REDACTED"
            } else {
                $affiliatedOrgNameNode = $affiliatedOrg.SelectSingleNode(".//a[contains(@class, 'value')]")
                $affiliatedOrgName = if ($affiliatedOrgNameNode) { $affiliatedOrgNameNode.InnerText.Trim() } else { "" }

                $affiliatedOrgSIDNode = $affiliatedOrg.SelectSingleNode(".//strong[contains(@class, 'value')]")
                $affiliatedOrgSID = if ($affiliatedOrgSIDNode) { $affiliatedOrgSIDNode.InnerText.Trim() } else { "" }

                $affiliatedOrgDetails += "$affiliatedOrgName($affiliatedOrgSID)"
            }
        }
    } else {
        $affiliatedOrgDetails += "None"
    }

    # Output the affiliated organizations in one line with special handling for "REDACTED"
    Write-Host -NoNewline "Affiliated Organizations: "

    foreach ($item in $affiliatedOrgDetails) {
        if ($item -eq "REDACTED") {
            Write-Host -NoNewline -ForegroundColor DarkRed $item
        } else {
            Write-Host -NoNewline -ForegroundColor DarkCyan $item
        }

        # Output a comma (,) after each organization, except for the last one
        if ($item -ne $affiliatedOrgDetails[-1]) {
            Write-Host -NoNewline ", "
        }
    }
}

# Check if the username is provided as a command-line argument
if ($args.Count -gt 0) {
    $username = $args[0]
} else {
    # Ask user for the username
    $username = Read-Host "Enter RSI Handle"
}

# URL for the RSI organizations page
$urlOrg = "https://robertsspaceindustries.com/citizens/$username/organizations"
# URL for the RSI citizen page
$urlCitizen = "https://robertsspaceindustries.com/citizens/$username"

# Use Invoke-WebRequest to get the HTML content of the pages
$htmlContentOrg = Invoke-WebRequest -Uri $urlOrg -UseBasicParsing | Select-Object -ExpandProperty Content
$htmlContentCitizen = Invoke-WebRequest -Uri $urlCitizen -UseBasicParsing | Select-Object -ExpandProperty Content

# Extract information using the HTML Agility Pack functions
$handleName = Extract-InfoWithHtmlAgilityPack -htmlContent $htmlContentCitizen -label "Handle name"
$ueeCitizenRecord = Extract-InfoWithHtmlAgilityPack -htmlContent $htmlContentCitizen -label "UEE Citizen Record"
$badge = Extract-BadgeInfoWithHtmlAgilityPack -htmlContent $htmlContentCitizen
$enlistedDate = Extract-EnlistedDateWithHtmlAgilityPack -htmlContent $htmlContentCitizen
$fluency = Extract-FluencyWithHtmlAgilityPack -htmlContent $htmlContentCitizen

# Output the extracted information from the citizen page
Write-Host -NoNewline ("Information extracted from: "); Write-Host -ForegroundColor DarkCyan $urlCitizen
Write-Host -NoNewline "Handle Name: "; Write-Host -ForegroundColor DarkCyan $handleName
Write-Host -NoNewline "UEE Citizen Record: "; Write-Host -ForegroundColor DarkCyan $ueeCitizenRecord
Write-Host -NoNewline "Badge: "; Write-Host -ForegroundColor DarkCyan $badge
Write-Host -NoNewline "Enlisted Date: "; Write-Host -ForegroundColor DarkCyan $enlistedDate
Write-Host -NoNewline "Fluency: "; Write-Host -ForegroundColor DarkCyan $fluency

# Call the function to get organization details
$orgDetails = Get-OrganizationDetails -htmlContent $htmlContentOrg

# Display the organization details
$orgDetails
