#Requires -Version 3.0
#The script requires the following permissions:
#    Application.Read.All (required)
#    AuditLog.Read.All (optional, needed to retrieve Sign-in stats)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5940/reporting-on-entra-id-application-registrations

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -Verbose
Param(
    [switch]$SkipExcelOutput=$false,
    [int]$ExcessiveIntervalInDays=180
)


#==========================================================================
#Helper functions
#==========================================================================

function Convert-AppPermissions {
    Param(
        [Parameter(Mandatory=$true)]$AppRoleAssignments
    )

    $appCount = 0
    $calendarAppCount = 0 
    $contactsAppCount = 0 
    $mailsAppCount = 0 
    $riskyAppCount = 0 
    $directoryAppCount = 0 
    $filesAppCount = 0 
    $sitesAppCount = 0 
    $readWriteAppCount = 0

    foreach ($AppRoleAssignment in $AppRoleAssignments) {
        $resID = (Get-ServicePrincipalRoleById $AppRoleAssignment.resourceAppId).appDisplayName
        foreach ($entry in $AppRoleAssignment.resourceAccess) {
            $entryValue = switch ($entry.Type) {
                "Role" { ($OAuthScopes[$AppRoleAssignment.resourceAppId].AppRoles | Where-Object {$_.id -eq $entry.id}).Value }
                "Scope" { ($OAuthScopes[$AppRoleAssignment.resourceAppId].publishedPermissionScopes | Where-Object {$_.id -eq $entry.id}).Value }
                default { continue }
            }
            if (!$entryValue) { $entryValue = "Orphaned ($($entry.id))" }
            $targetPermissions = if ($entry.Type -eq "Role") { $OAuthpermA } else { $OAuthpermD }
            $targetPermissions["[$resID]"] += "," + $entryValue
        }
        if ($entryValue) {
            $scopes = $entryValue.Split(" ")
            $calendarAppCount += ($scopes -like "*Calendars*").Count
            $contactsAppCount += ($scopes -like "*Contacts*").Count
            $mailsAppCount += ($scopes -like "*Mail.*").Count
            $riskyAppCount += ($scopes -like "AppRoleAssignment.ReadWrite.All").Count
            $directoryAppCount += ($scopes -like "Directory.ReadWrite*").Count
            $filesAppCount += ($scopes -like "Files*").Count
            $sitesAppCount += ($scopes -like "Sites*").Count
            $readWriteAppCount += ($scopes -like "*ReadWrite*").Count
            $appCount++ 
        }
    }
    return $appCount, $calendarAppCount, $contactsAppCount, $mailsAppCount, $riskyAppCount, $directoryAppCount, $filesAppCount, $sitesAppCount, $readWriteAppCount
}

function Get-ServicePrincipalRoleById {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$resID)

    #check if we've already collected this SP data
    if (!$OAuthScopes[$resID]) {
        $OAuthScopes[$resID] = Get-MgBetaServicePrincipal -Filter "appId eq '$resID'" -ErrorAction Stop -Verbose:$false
    }
    return $OAuthScopes[$resID]
}

function Convert-Credential {
    Param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$cred,
        [bool]$ComputeScore = $false
    )

    $credout = @($null,@())
    $credout[0] = ($cred.count).ToString()
    if ((Get-Date) -gt ($cred.endDateTime |  Sort-Object -Descending | Select-Object -First 1)) { $credout[0] += " (expired)" }
    foreach ($c in $cred) {
        $cstring = $c.keyId
        if ((New-TimeSpan -Start $c.startDateTime -End $c.endDateTime).Days -ge $ExcessiveIntervalInDays) { 
            $excessiveValidity = $true
            if ($ComputeScore) {
                $script:appWithExcessiveValidity++ 
                $script:appRiskScore++
            }
        }
        if ((Get-Date) -gt ($c.endDateTime)) { 
            $cstring += "(EXPIRED)" 
            if ($ComputeScore) {
                $script:appWithExpiredCreds++ 
                $script:appRiskScore++
            }
        }
        $cstring += "(valid from $($c.startDateTime) to $($c.endDateTime))"
        $credout[1] += $cstring
    }
    if ($excessiveValidity) { $credout[0] += " (excessive validity)" }

    return $credout
}

function Convert-SPSignInStats {

    Param(
        #Report object
        [Parameter(Mandatory=$true)]$SPSignInStats)
        
    foreach ($SPSignInStat in $SPSignInStats) {
        if (!$SPStats[$SPSignInStat.appId]) {
            $SPStats[$SPSignInStat.appId] = @{
                "LastSignIn" = $SPSignInStat.lastSignInActivity.lastSignInDateTime
                "LastDelegateClientSignIn" = $SPSignInStat.delegatedClientSignInActivity.lastSignInDateTime
                "LastDelegateResourceSignIn" = $SPSignInStat.delegatedResourceSignInActivity.lastSignInDateTime
                "LastAppClientSignIn" = $SPSignInStat.applicationAuthenticationClientSignInActivity.lastSignInDateTime
                "LastAppResourceSignIn" = $SPSignInStat.applicationAuthenticationResourceSignInActivity.lastSignInDateTime
                "LastActivityDate" = @(
                    $SPSignInStat.lastSignInActivity.lastSignInDateTime,
                    $SPSignInStat.delegatedClientSignInActivity.lastSignInDateTime,
                    $SPSignInStat.delegatedResourceSignInActivity.lastSignInDateTime,
                    $SPSignInStat.applicationAuthenticationClientSignInActivity.lastSignInDateTime,
                    $SPSignInStat.applicationAuthenticationResourceSignInActivity.lastSignInDateTime
                ) | Sort-Object -Descending | Select-Object -First 1
            }
        }
    }
    #return $SPStats
}

function Convert-AppCredStats {

    Param(
            [Parameter(Mandatory=$true)]$AppCredStats
        )

    foreach ($AppCredStat in $AppCredStats) {
        if (!$AppCreds[$AppCredStat.appId]) {
            $AppCreds[$AppCredStat.appId] = @{
                "LastSignIn" = $AppCredStat.signInActivity.lastSignInDateTime
            }
        }
    }
}

function Convert-ToHtml {
    param(
        [Parameter(Mandatory=$true)]$csvContent
    )
    
    $style = @"
    <style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #18293b;
        background-color: #eef1f5;
        background-image: url(https://cdn.prod.website-files.com/612933c2d902f2ac80205a6f/65b144686594df863fb6249f_Purple-gradient-background-web.svg);
        padding: 10px;
        margin: 0;
    }
    .flex-container {
        display: flex;
        justify-content: space-between;
    }
    .progress-container {
        display: flex;
        align-items: center;
    }
    .center{
        text-align: center;
    }
    .text-container {
        flex: 1;
        padding: 0px 10px;
    }
    .box{
        align-items: top;
        padding: 10px 20px;             
        font-family: 'Arial', sans-serif; 
        color: #18293b;               
        width: auto;   
        margin: 10px;
        min-width: 30%;
    }
    .white{
        border-radius: 8px;        
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        background-color: #ffffff; 
    }
    .right{
        width: 70%;
    }
    .left{
        width: 30%;
    }
    .gradient-button {
        background: linear-gradient(90deg, #A60066 0%, #3C67BF 100%);
        border: none;
        border-radius: 8px;
        color: white;
        padding: 14px 24px;
        font-size: 16px;
        font-weight: 600;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        cursor: pointer;
        transition: transform 0.3s;
    }

    .gradient-button:hover {
        transform: scale(1.05);
    }
    .tooltip {
        position: relative;
        display: table;
        border-bottom: 1px dotted #18293b; /* Visual hint for the tooltip */
    }
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 200px;
        background-color: #18293b;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 5px 0;
        position: absolute;
        z-index: 1;
        bottom: 125%;
        left: 50%;
        margin-left: -100px;
        opacity: 0;
        transition: opacity 0.3s;
    }
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
    .logo {
        display: block;
        margin: 0 auto 20px auto;
        position: absolute;
        top: 25px;
        right: 20px;
        height: 50px;
        max-width: 40%;
    }
    .critical
    {
        color: red !important;
        font-weight: bold;
        font-size: 1.1em;
    }
    .medium
    {
        font-weight: bold;
        font-size: 1.1em;
    }
    h1 {
        padding: 0px 10px;
    }
    h1, h3 {
        max-width: 60%;
        color: #18293b;
    }
    span, li, td, th {
        color: #18293b;
        font-size: 0.9em;
    }
    ul {
        list-style-type: none;
        padding: 0px;
    }
    li {
        margin-bottom: 10px;
    }
    th, td {
        padding: 2px; 
        overflow-wrap: break-word;
        white-space: nowrap;
    }
    tr:hover {
        background-color: #f5f5f5;
    }
    tr:nth-child(even) {
        background-color: #eef1f5;
    }
    .progress-circle {
        width: 100px;
        height: 100px;
        border-radius: 50%;
        display: flex;
        justify-content: center;
        align-items: center;
        background: var(--progress, white); /* Fallback to white if --progress is not set */
        position: relative;
        --progress: conic-gradient(white 0deg, white 360deg); /* Default value */
    }

    .progress-circle::before {
        content: '';
        position: absolute;
        width: 90px; /* Smaller than the outer circle */
        height: 90px; /* Smaller than the outer circle */
        background: #f5f5f5; /* Same as the body background */
        border-radius: 50%;
        z-index: 1; /* Ensure it sits above the pseudo element for the progress */
    }

    .grade-letter {
        position: absolute;
        font-size: 3em;
        font-weight: bold;
        z-index: 2; /* Ensure the text sits above the inner circle */
    }
    .table-container {
        overflow: auto;
        width: 100%; /* Enable scrolling */
        max-height: 800px; 
        margin: 10px;
    }
    table {
        background-color: #fff;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); /* Light shadow */
        border-radius: 6px;
        padding: 20px;
        border-radius: 8px; 
        font-size: 0.9em;
        width: 100%; /* Ensure the table takes full width */
        table-layout: auto;
        text-align: left;
        width: auto; /* Ensure the table takes full width */
        border-collapse: collapse; /* Ensure borders collapse */
    }
    th, td {
        padding: 8px; /* Add padding for better readability */
        text-align: left; /* Align text to the left */
        border: 1px solid #ddd; /* Add border to cells */
    }
    /* Sticky header */
    th {
        position: sticky;
        top: 0;
        z-index: 1; /* Ensure header is above other content */
        background-color: #f2f2f2; /* Background color to cover underlying content */
    }
    th, th:first-child{
        position: sticky;
        top: 0;
        z-index: 2; /* Ensure header is above other content */
        background-color: #f2f2f2; /* Background color to cover underlying content */
    }
    /* Optional: Sticky first column */
    th, td:first-child {
        position: sticky;
        left: 0;
        z-index: 1; /* Lower z-index to keep behind the header */
        background-color: #e6e6e6; /* Light grey background */
    }
    tr:nth-child(even) {
        background-color: #f2f2f2; /* Add background color to even rows */
    }
    tr:hover {
        background-color: #ddd; /* Add hover effect */
    }
    </style>
"@

# Logic to compute the grade and text
$grade = if ($script:currentScore -eq 0) { 0 } else { ($script:currentScore / $script:totalScore) * 100 }
$ctaText = if ($grade -gt 89) {
    "<b>Keep monitoring your organization's defenses.</b>"
} elseif ($grade -lt 68) {
    "<b>Take the next step in fortifying your defenses, now.</b>"
} else {
    "<b>Your organization has some issues with app registrations.</b>"
}
$ctaText += " Future-proof app security. Automate best practice security across your <i>entire</i> Microsoft 365 environment--with CoreView. See the platform in action."
$gradeText = if ($grade -gt 89) {
    "<b>You are amazing!</b> We manage more than 1500 customers worldwide and you are sure in the top tier!"
} elseif ($grade -lt 68) {
    "<b>Your organization has risky app registrations.</b> We strongly suggest you to review results in the table or in the excel produced and act now."
} else {
    "<b>Your organization has some issues with app registrations.</b> Internal developed apps can be a risk if not constantly reviewed. Please check table data and fix as many issues as possible."
}

    $head = @"
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    $style
    <title>Entra App Registrations</title>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script>
      google.charts.load('current', {'packages':['corechart']});
      function drawChart() {
        var currentScore = $script:currentScore;
        var riskScore = $script:totalScore - $script:currentScore;

        var data = google.visualization.arrayToDataTable([
          // Add a header row for labels and each series
          ['Score', 'Percentage'],
          ['Healthy Score', currentScore],
          ['Risky Score', riskScore]
        ]);
        var colors = ['#267530', '#9F0D1B'];
        var options = {
            colors: colors,
            backgroundColor: {
                fill: '#eef1f5', 
                fillOpacity: 0 
            },
            chartArea: {
                width: '80%', 
                height: '80%'
            }
        };
        var chart = new google.visualization.PieChart(document.getElementById('piechart'));
        chart.draw(data, options);
        }
        google.charts.setOnLoadCallback(drawChart);
    </script>
    <script>
        function updateProgress(percentage) {
        const circle = document.getElementById('progressCircle');
        const gradeLetter = document.getElementById('gradeLetter');

        const gradeScale = [
            { min: 97, letter: 'A+', color: 'green' },
            { min: 93, letter: 'A', color: 'green' },
            { min: 90, letter: 'A-', color: 'green' },
            { min: 87, letter: 'B+', color: 'green' },
            { min: 83, letter: 'B', color: 'green' },
            { min: 80, letter: 'B-', color: 'green' },
            { min: 77, letter: 'C+', color: 'orange' },
            { min: 73, letter: 'C', color: 'orange' },
            { min: 70, letter: 'C-', color: 'orange' },
            { min: 67, letter: 'D+', color: 'red' }, 
            { min: 63, letter: 'D', color: 'red' },
            // No 'E' grade in the US grading system
            { min: 0,  letter: 'F', color: 'red' }
        ];

        const grade = gradeScale.find(g => percentage >= g.min);
        gradeLetter.textContent = grade.letter;
        gradeLetter.style.color = grade.color; 

        const angle = percentage * 3.6;
        circle.style.setProperty("--progress", ``conic-gradient(`${grade.color} `${angle}deg, #f5f5f5 `${angle}deg)``);
        }
    </script>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            updateProgress($grade);
        });
    </script>
"@

    $body = @"
    <a href="https://www.coreview.com/pricing">
        <img src="https://cdn.prod.website-files.com/612933c2d902f2ac80205a6f/65b28b20723fca6b6bb16213_CoreView-dark-logo.svg" alt="CoreView SimeonCloud Logo" class="logo" />
    </a>
    <h1>Entra Security Scanner for App Registrations</h1>
    <div class='flex-container'>
        <div class="white box left">
            <div class="progress-container">
                <div class="progress-circle" id="progressCircle">
                    <div class="grade-letter" id="gradeLetter">$grade</div>
                </div>
                <div class="text-container">
                    <p>$gradeText</p>
                </div>
            </div>
        </div>
        <div class="white box right center">
            <div>
                <p>$ctaText</p>
                <a href='https://www.coreview.com/request-a-demo?utm_medium=Platform&utm_source=popup&utm_campaign=2024-Q2-WW-CM-DEMO-REQUEST-FT:%20Entra%20Security%20Scanner%20for%20Custom%20AppRegistration' target='_blank' class="gradient-button">Book a demo</a>
            </div>
        </div>
    </div>
    <div class='flex-container'>
        <div class="box left">
            <h3>Relevant info:</h3>
            <ul>
                <li class="tooltip">Report made on: $(Get-Date)</li>
                <li class="tooltip">Apps analyzed: $script:totalApps</li>
                <li class="tooltip">Apps created in the last 30 days: $appsCreatedLast30Days</li>
                <br>
                <li class="critical"><b>Critical issues</b>
                <li class="tooltip"><b>Apps without owners: $script:appsWithoutOwners</b><span class="tooltiptext">These apps have now owners</span></li>
                <li class="tooltip"><b>Apps with risky access (Midnight Blizzard attack vector): $appsWithRiskyAccess</b><span class="tooltiptext">These apps can create others apps with any consent. Search for 'AppRole' to find apps</span></li>
                <li class="tooltip"><b>Bad URIs: $script:appWithBadURIs</b><span class="tooltiptext">These apps have bad URIs containing localhost or http:// or any (*)</span></li>
                <br>
                <li class="medium"><b>Medium issues</b>
                <li class="tooltip">Expired Credentials: $script:appWithExpiredCreds<span class="tooltiptext">These apps have at least 1 expired credential. Search for 'expired' to find apps</span></li>
                <li class="tooltip">Excessive Validity ($ExcessiveIntervalInDays days): $script:appWithExcessiveValidity<span class="tooltiptext">These apps have excessive validity. Search for 'excessive validity' to find apps</span></li>
                <li class="tooltip">Unused Apps ($ExcessiveIntervalInDays days): $script:unusedApps<span class="tooltiptext">These apps are not being used in the last $ExcessiveIntervalInDays days</span></li>
                <br>
                <li class="tooltip"><b>Info</b>
                <li class="tooltip">Apps with Directory.ReadWrite access: $appsWithDirectoryReadWriteAccess<span class="tooltiptext">These apps can access and modify any Directory resources. Search for 'Directory.ReadWrite' to find apps</span></li>
                <li class="tooltip">Apps with ReadWrite access: $appsWithReadWriteAccess<span class="tooltiptext">These apps can access and modify any resources. Search for 'ReadWrite' to find apps</span></li>
                <li class="tooltip">Apps with calendars access: $appsWithCalendarAccess<span class="tooltiptext">These apps can access and/or modify any calendar information. Search for 'Calendars.' to find apps</span></li>
                <li class="tooltip">Apps with contacts access: $appsWithContactsAccess<span class="tooltiptext">These apps can access and/or modify any contacts information. Search for 'Contacts.' to find apps</span></li>
                <li class="tooltip">Apps with mails access: $appsWithMailAccess<span class="tooltiptext">These apps can access and/or modify any mailbox. Search for 'Mail' to find apps</span></li>
                <li class="tooltip">Apps with files access: $appsWithFilesAccess<span class="tooltiptext">These apps can access and/or modify any OneDrive. Search for 'Files.' to find apps</span></li>
                <li class="tooltip">Apps with sites access: $appsWithSitesAccess<span class="tooltiptext">These apps can access and/or modify SharePoint site. Search for 'Sites.' to find apps</span></li>
            </ul>
        </div>
        <div id="piechart" class="box right"></div>
    </div>
"@

    $htmlContent = $csvContent | ConvertTo-Html -Head $head -Body $body -PreContent "<div class='table-container'>" -PostContent "</div>"
    return $htmlContent
}

#==========================================================================
#Main script starts here
#==========================================================================

$RequiredScopes = switch ($PSBoundParameters.Keys) {
    Default { "Application.Read.All", "AuditLog.Read.All" }
}

Write-Verbose "Connecting to Graph API..."
Import-Module Microsoft.Graph.Beta.Applications -Verbose:$false -ErrorAction Stop
try {
    Connect-MgGraph -Scopes $RequiredScopes -verbose:$false -ErrorAction Stop -NoWelcome
}
catch { throw $_ }

$CurrentScopes = (Get-MgContext).Scopes
if ($RequiredScopes | Where-Object {$_ -notin $CurrentScopes }) { Write-Error "The access token does not have the required permissions, rerun the script and consent to the missing scopes!" -ErrorAction Stop }

$Apps = @()
#Prepare variables
$OAuthScopes = @{} #hash-table to store data for app roles and stuff
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0
$script:totalScore = 0
$script:appRiskScore = 0
$script:currentScore = 0
$script:totalApps = 0
$appsWithAccess = 0
$appsWithCalendarAccess = 0
$appsWithContactsAccess = 0
$appsWithMailAccess = 0
$appsWithRiskyAccess = 0
$appsCreatedLast30Days = 0
$appsWithDirectoryReadWriteAccess = 0
$appsWithReadWriteAccess = 0
$appsWithFilesAccess = 0
$appsWithSitesAccess = 0
$createdInInterval = (Get-Date).AddDays(-30)
$script:appWithAdal = 0
$script:appWithExpiredCreds = 0
$script:unusedApps = 0
$script:appWithUnusedCreds = 0
$script:appWithBadURIs = 0
$script:appWithExcessiveValidity = 0
$script:appsWithoutOwners = 0

Write-Verbose "Retrieving list of applications..."
$Apps = Get-MgBetaApplication -All -ErrorAction Stop -Verbose:$false

Write-Verbose "Retrieving sign-in stats for service principals..."
$SPSignInStats = Get-MgBetaReportServicePrincipalSignInActivity -All -ErrorAction Stop -Verbose:$false
$SPStats = @{} 
if ($SPSignInStats) { Convert-SPSignInStats $SPSignInStats }

Write-Verbose "Retrieving application credential usage stats..."
$AppCredStats = Get-MgBetaReportAppCredentialSignInActivity -All -ErrorAction Stop -Verbose:$false
$AppCreds = @{} 
if ($AppCredStats) { Convert-AppCredStats $AppCredStats }

foreach ($App in $Apps) {
    $script:totalApps = @($Apps).count
    $script:totalScore += 10
    $script:currentScore += 10
    $script:appRiskScore = 0
    $ActivityMessage = "Retrieving data for application $($App.DisplayName). Please wait..."
    $StatusMessage = ("Processing application {0} of {1}: {2}" -f $count, @($Apps).count, $App.id)
    $PercentComplete = ($count / @($Apps).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing application $($App.id)..."

    Write-Verbose "Retrieving owners info..."
    $owners = @()
    $owners = Get-MgBetaApplicationOwner -ApplicationId $App.id -Property id,userPrincipalName -All -ErrorAction Stop -Verbose:$false
    if ($owners) { $owners = $owners.AdditionalProperties.userPrincipalName }
    else { $script:appsWithoutOwners++; $script:appRiskScore+=2}

    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = (&{if ($App.DisplayName) { $App.DisplayName } else { $null }})
        "ApplicationId" = $App.AppId
        "Publisher Domain" = (&{if ($App.PublisherDomain) { $App.PublisherDomain } else { $null }})
        "Verified" = (&{if ($App.verifiedPublisher.verifiedPublisherId) { $App.verifiedPublisher.displayName } else { "Not verified" }})
        "Certification" = (&{if ($App.certification) { $App.certification.certificationDetailsUrl } else { "" }})
        "SignInAudience" = $App.signInAudience
        "ObjectId" = $App.id
        "Created on" = (&{if ($App.createdDateTime) { (Get-Date($App.createdDateTime) -format g) } else { "N/A" }})
        "Owners" = (&{if ($owners) { $owners -join "," } else { $null }})
        "Permissions (application)" = $null
        "Permissions (delegate)" = $null
        "Permissions (API)" = $null
        "Allow Public client flows" = (&{if ($App.isFallbackPublicClient -eq "true") { "True" } else { "False" }}) 
        "Key credentials" = (&{if ($App.keyCredentials) { (Convert-Credential $App.keyCredentials $true)[0] } else { "" }})
        "KeyCreds" = (&{if ($App.keyCredentials) { ((Convert-Credential $App.keyCredentials)[1]) -join ";" } else { $null }})
        "Next expiry date (key)" = (&{if ($App.keyCredentials) { ($App.keyCredentials.endDateTime | Where-Object {$_ -ge (Get-Date)} |  Sort-Object -Descending | Select-Object -First 1) } else { "" }})
        "Password credentials" = (&{if ($App.passwordCredentials) { (Convert-Credential $App.passwordCredentials $true)[0] } else { "" }})
        "PasswordCreds" = (&{if ($App.passwordCredentials) { ((Convert-Credential $App.passwordCredentials)[1]) -join ";" } else { $null }})
        "Next expiry date (password)" = (&{if ($App.passwordCredentials) { ($App.passwordCredentials.endDateTime | Where-Object {$_ -ge (Get-Date)} |  Sort-Object -Descending | Select-Object -First 1) } else { "" }})
        "App property lock" = (&{if ($App.servicePrincipalLockConfiguration.isEnabled -and $App.servicePrincipalLockConfiguration.allProperties) { $true } else { $false }})
        "HasBadURIs" = (&{if ($App.web.redirectUris -match "localhost|http://|urn:|\*") { $true; $script:appWithBadURIs++; $script:appRiskScore+=2} else { $false }})
        "Redirect URIs" = (&{if ($App.web.redirectUris) { $App.web.redirectUris -join ";" } else { $null }})
        "Total consents" = $null
        "Calendar" = $null
        "Contacts" = $null
        "Mails" = $null
        "Directory" = $null
        "Risky" = $null
        "Files" = $null
        "Sites" = $null
        "ReadWrite" = $null
    }

    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last sign-in" -Value (&{if ($SPStats[$App.appId].LastSignIn) { (Get-Date($SPStats[$App.appid].LastSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate client sign-in" -Value (&{if ($SPStats[$App.appid].LastDelegateClientSignIn) { (Get-Date($SPStats[$App.appid].LastDelegateClientSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last delegate resource sign-in" -Value (&{if ($SPStats[$App.appid].LastDelegateResourceSignIn) { (Get-Date($SPStats[$App.appid].LastDelegateResourceSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app client sign-in" -Value (&{if ($SPStats[$App.appid].LastAppClientSignIn) { (Get-Date($SPStats[$App.appid].LastAppClientSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last app resource sign-in" -Value (&{if ($SPStats[$App.appid].LastAppResourceSignIn) { (Get-Date($SPStats[$App.appid].LastAppResourceSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last credential sign-in" -Value (&{if ($AppCreds[$App.appid].LastSignIn) { (Get-Date($AppCreds[$App.appid].LastSignIn) -format g) } else { $null }})
    $objPermissions | Add-Member -MemberType NoteProperty -Name "Last activity date" -Value (&{if ($SPStats[$App.appid].LastActivityDate) { (Get-Date($SPStats[$App.appid].LastActivityDate) -format g) } else { $null }})
    if (!$objPermissions."Last activity date" -or (New-TimeSpan -Start $objPermissions."Last activity date" -End (Get-Date)).Days -ge $ExcessiveIntervalInDays) {
        $script:unusedApps++
        $script:appRiskScore++;
    }

    if ($App.requiredResourceAccess | Where-Object {$_.resourceAppId -eq "00000002-0000-0000-c000-000000000000"}) {
        $objPermissions | Add-Member -MemberType NoteProperty -Name "UsesAADGraph" -Value $true
    }
    else { $objPermissions | Add-Member -MemberType NoteProperty -Name "UsesAADGraph" -Value $false }

    $OAuthpermA = @{};$OAuthpermD = @{};$resID = $null;

    if ($App.requiredResourceAccess) { 
        $appsPermissionsCounts = Convert-AppPermissions $App.requiredResourceAccess
        $objPermissions.'Total consents' = $appsPermissionsCounts[0]
        $objPermissions.'Calendar' = $appsPermissionsCounts[1]
        $objPermissions.'Contacts' = $appsPermissionsCounts[2]
        $objPermissions.'Mails' = $appsPermissionsCounts[3]
        $objPermissions.'Risky' = $appsPermissionsCounts[4]
        $objPermissions.'Directory' = $appsPermissionsCounts[5]
        $objPermissions.'Files' = $appsPermissionsCounts[6] 
        $objPermissions.'Sites' = $appsPermissionsCounts[7] 
        $objPermissions.'ReadWrite' = $appsPermissionsCounts[8] 
        if ($appsPermissionsCounts[0] -gt 0) {$appsWithAccess++}
        if ($appsPermissionsCounts[1] -gt 0) {$appsWithCalendarAccess++}
        if ($appsPermissionsCounts[2] -gt 0) {$appsWithContactsAccess++}
        if ($appsPermissionsCounts[3] -gt 0) {$appsWithMailAccess++;}
        if ($appsPermissionsCounts[4] -gt 0) {$appsWithRiskyAccess++; $script:appRiskScore+=2}
        if ($appsPermissionsCounts[5] -gt 0) {$appsWithDirectoryReadWriteAccess++}
        if ($appsPermissionsCounts[6] -gt 0) {$appsWithFilesAccess++}
        if ($appsPermissionsCounts[7] -gt 0) {$appsWithSitesAccess++}
        if ($appsPermissionsCounts[8] -gt 0) {$appsWithReadWriteAccess++}
    }
    else { 
        Write-Verbose "No permissions found for application $($App.id), skipping..."
        $objPermissions.'Total consents' = 0
        $objPermissions.'Calendar' = 0
        $objPermissions.'Contacts' = 0
        $objPermissions.'Mails' = 0
        $objPermissions.'Risky' = 0
        $objPermissions.'Directory' = 0
        $objPermissions.'Files' = 0
        $objPermissions.'Sites' = 0
        $objPermissions.'ReadWrite' = 0
    }

    if ($App.CreatedDateTime -and ((Get-Date($App.CreatedDateTime)) -gt $createdInInterval)) {
        $appsCreatedLast30Days++
    }

    $objPermissions.'Permissions (application)' = (($OAuthpermA.GetEnumerator() | ForEach-Object { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    $objPermissions.'Permissions (delegate)' = (($OAuthpermD.GetEnumerator() | ForEach-Object { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join ";")
    if ($App.api) { $objPermissions.'Permissions (API)' = ($App.api.oauth2PermissionScopes.value -join ";") }
    if ($script:appRiskScore -gt 10) { $script:currentScore -= 10 } else { $script:currentScore -= $script:appRiskScore }
    $output.Add($objPermissions)
}

$output = $output | Select-Object 'Application Name','Publisher Domain','Verified', 'SignInAudience', `
'Created on', 'Owners', "Permissions (application)",  "Permissions (delegate)",  "Permissions (api)", `
'Key credentials', 'Next expiry date (key)','Password credentials', 'Next expiry date (password)', `
"Total consents", "Calendar","Contacts","Mails", 'Directory', "Risky", "Files", "Sites", "ReadWrite", `
"Last activity date","Last sign-in","Last delegate client sign-in",  "Last delegate resource sign-in", `
"Last app client sign-in", "Last app resource sign-in","Last credential sign-in", `
'HasBadURIs', 'Redirect URIs', 'App property lock', 'Certification',  "Allow Public client flows", 'KeyCreds', `
'ApplicationId', 'ObjectId', 'UsesAADGraph', 'PasswordCreds' -ExcludeProperty Number 
$outputFilePath = "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory"
$output | Export-CSV -nti -Path "$outputFilePath.csv"
Write-Host "Output exported to $($PWD)\$($outputFilePath).csv"
$csvContent = Import-Csv "$outputFilePath.csv"

$htmlContent = Convert-ToHtml -csvContent $csvContent
$htmlAppsOutputPath = "$($outputFilePath)_apps.html"
$htmlContent | Out-File -Encoding utf8 $htmlAppsOutputPath

if (-not $CsvOnly)
{
    Write-Host "Output exported to $($($outputFilePath)).xlsx"
    $output | Select-Object * -ExcludeProperty Number | Export-Excel -Path "$($outputFilePath).xlsx" -AutoSize -TableName "AppRegistrations"
}