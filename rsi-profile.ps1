# Hardcoded path to the HTML Agility Pack DLL
$htmlAgilityPackPath = "PATH TO YOUR DLL"

# Load the HTML Agility Pack assembly
Add-Type -Path $htmlAgilityPackPath

# Function to create an HTML Agility Pack document
function New-HtmlDocument {
    param (
        [string]$htmlContent
    )

    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($htmlContent)

    return $doc
}

# Function to extract information from HTML using HTML Agility Pack
function Extract-InfoWithHtmlAgilityPack {
    param (
        [string]$htmlContent,
        [string]$label
    )

    $doc = New-HtmlDocument -htmlContent $htmlContent

    # Create an HTML Agility Pack document
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

    $doc = New-HtmlDocument -htmlContent $htmlContent

    # Create an HTML Agility Pack document
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

    $doc = New-HtmlDocument -htmlContent $htmlContent

    # Create an HTML Agility Pack document
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

    $doc = New-HtmlDocument -htmlContent $htmlContent

    # Create an HTML Agility Pack document
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

    $doc = New-HtmlDocument -htmlContent $htmlContent

    # Use HTML Agility Pack to load the HTML content
    $mainOrg = $doc.DocumentNode.SelectSingleNode("//div[contains(@class, 'box-content org main')]")
    if ($mainOrg) {
        $visibilityClass = $mainOrg.GetAttributeValue("class", "")
        $isRedacted = $visibilityClass -like "*visibility-R*"

        if ($isRedacted) {
            Write-Host -NoNewline -ForegroundColor DarkYellow "Main Organization: "; Write-Host -ForegroundColor DarkRed "REDACTED"
        } else {
            $mainOrgNameNode = $mainOrg.SelectSingleNode(".//a[@class='value']")
            $mainOrgName = if ($mainOrgNameNode) { $mainOrgNameNode.InnerText.Trim() } else { "" }

            $mainOrgSIDNode = $mainOrg.SelectSingleNode(".//strong[@class='value']")
            $mainOrgSID = if ($mainOrgSIDNode) { $mainOrgSIDNode.InnerText.Trim() } else { "" }

            Write-Host -NoNewline -ForegroundColor DarkYellow "Main Organization: "; Write-Host -ForegroundColor DarkCyan "$mainOrgName($mainOrgSID)"
        }
    } else {
        Write-Host -NoNewline -ForegroundColor DarkYellow "Main Organization: "; Write-Host -ForegroundColor DarkCyan "None"
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
    Write-Host -NoNewline -ForegroundColor DarkYellow "Affiliated Organizations: "

    $isFirst = $true

    foreach ($item in $affiliatedOrgDetails) {
        if ($item -eq "REDACTED") {
            if (-not $isFirst) {
                Write-Host -NoNewline -ForegroundColor DarkYellow " | "
            }
            Write-Host -NoNewline -ForegroundColor DarkRed $item
        } else {
            if (-not $isFirst) {
                Write-Host -NoNewline -ForegroundColor DarkYellow " | "
            }
            Write-Host -NoNewline -ForegroundColor DarkCyan $item
        }

        $isFirst = $false
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
try {
    $htmlContentOrg = Invoke-WebRequest -Uri $urlOrg -UseBasicParsing -ErrorAction Stop | Select-Object -ExpandProperty Content
    $htmlContentCitizen = Invoke-WebRequest -Uri $urlCitizen -UseBasicParsing -ErrorAction Stop | Select-Object -ExpandProperty Content
}
catch {
    $errorResponse = $_.Exception.Response

    if ($errorResponse -eq $null) {
        Write-Host "An unexpected error occurred: $_"
        exit
    }

    if ($errorResponse.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
        Write-Host "Citizen '$username' was not found (404). Try again."
        exit
    }

    Write-Host "HTTP Error: $($errorResponse.StatusCode.value__)"
    Write-Host "Response: $($errorResponse.StatusDescription)"
    exit
}

function Get-UserDetails {
    param (
        [string]$htmlContent
    )

    # Extract information using the HTML Agility Pack functions
    $handleName = Extract-InfoWithHtmlAgilityPack -htmlContent $htmlContent -label "Handle name"
    $ueeCitizenRecord = Extract-InfoWithHtmlAgilityPack -htmlContent $htmlContent -label "UEE Citizen Record"
    $badge = Extract-BadgeInfoWithHtmlAgilityPack -htmlContent $htmlContent
    $enlistedDate = Extract-EnlistedDateWithHtmlAgilityPack -htmlContent $htmlContent
    $fluency = Extract-FluencyWithHtmlAgilityPack -htmlContent $htmlContent

    # Output the extracted information from the citizen page
    Write-Host -NoNewline -ForegroundColor DarkGray ("Information extracted from: "); Write-Host -ForegroundColor White $urlCitizen
    Write-Host -NoNewline -ForegroundColor DarkYellow "Handle Name: "; Write-Host -ForegroundColor DarkCyan $handleName
    Write-Host -NoNewline -ForegroundColor DarkYellow "UEE Citizen Record: "; Write-Host -ForegroundColor DarkCyan $ueeCitizenRecord
    Write-Host -NoNewline -ForegroundColor DarkYellow "Badge: "; Write-Host -ForegroundColor DarkCyan $badge
    Write-Host -NoNewline -ForegroundColor DarkYellow "Enlisted Date: "; Write-Host -ForegroundColor DarkCyan $enlistedDate
    Write-Host -NoNewline -ForegroundColor DarkYellow "Fluency: "; Write-Host -ForegroundColor DarkCyan $fluency
}

# Print user details
$userDetails = Get-UserDetails -htmlContent $htmlContentCitizen
# Print organization details
$orgDetails = Get-OrganizationDetails -htmlContent $htmlContentOrg
