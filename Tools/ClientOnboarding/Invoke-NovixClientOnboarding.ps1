<#
.SYNOPSIS
    Onboards a client tenant: creates the GDAP relationship and role-mapped security
    groups in the Novix partner tenant, then applies Novix support branding inside the
    client tenant.

.DESCRIPTION
    Runs in two phases.

    Phase 1 - Partner tenant (Novix)
      * Ensures one security group per GDAP role exists ("M365 GDAP <Role>").
      * Optionally adds named users to every group.
      * Creates a delegated admin (GDAP) relationship for the customer, locks it for
        approval, prints/copies the approval URL, and waits for the customer to accept.
      * Creates one access assignment per role/group pair and waits for them to activate.

    Phase 2 - Customer tenant (delegated, via the GDAP relationship just created)
      * Sets the Entra ID sign-in page text (company branding).
      * Sets the Intune Company Portal default branding profile support information.
      * Prints the Microsoft 365 "Help desk information" values for manual entry -
        see the NOTES section for why this one is not automated.

    Every write is idempotent, so the script is safe to re-run. Use -WhatIf to preview.

.PARAMETER CustomerTenantId
    Customer tenant GUID.

.PARAMETER GdapGroupMember
    UPNs (in the Novix partner tenant) to add to every GDAP role group.

.PARAMETER SkipGdap
    Skip phase 1. Use this to apply branding to a client that is already onboarded.

.PARAMETER SkipBranding
    Skip phase 2 entirely (GDAP setup only).

.PARAMETER SkipSignInBranding
    Skip only the Entra sign-in page text.

.PARAMETER SkipIntuneBranding
    Skip only the Intune Company Portal branding profile.

.PARAMETER ApplyIntuneVisualBranding
    Also set the Company Portal visual identity (organization name, Novix theme colour,
    and "show organization name only" in the header). Off by default because it
    overwrites the client's own Company Portal look, including hiding their logo.
    Support information is set either way.

.PARAMETER IntuneOrganizationName
    Organization name shown in the Company Portal, used only with
    -ApplyIntuneVisualBranding. Defaults to the customer tenant's own display name,
    which is almost always what you want in a client tenant.

.PARAMETER RelationshipDurationIso8601
    GDAP relationship duration. Maximum supported by Microsoft is P2Y.

.PARAMETER Force
    Skip the "you are about to write to <tenant>" confirmation in phase 2.

.EXAMPLE
    .\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '00000000-1111-2222-3333-444444444444'

    Full onboarding: GDAP groups, relationship, approval wait, then branding.

.EXAMPLE
    .\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -SkipGdap

    Apply/refresh Novix branding on a client that already has an active GDAP relationship.

.EXAMPLE
    .\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -SkipGdap -WhatIf

    Show exactly what would change in the client tenant without writing anything.

.NOTES
    Sign-in page text
      Entra ID caps signInPageText at 1024 characters and renders a limited markdown
      subset (bold, italics, and [label](url) hyperlinks). The Novix text is 1022
      characters when joined with LF. It is built here from an array joined with "`n"
      rather than a here-string on purpose: a here-string picks up the file's CRLF line
      endings, which pushes the same text to 1036 characters and the API call fails
      validation. Keep the array-join if you edit the wording, and keep an eye on the
      length guard below.

    Microsoft 365 help desk information
      Microsoft exposes no supported Graph API for Settings > Org settings >
      Organization profile > Help desk information (CIPP has no standard for it either,
      for the same reason). The only programmatic route is the undocumented
      admin.microsoft.com portal API, which needs a token for a non-Graph audience and
      can change without notice. Rather than ship a guess, phase 2 prints the exact
      values and the portal link so it is a 30-second paste. If we capture the real
      request from the browser dev tools when saving that form, it can be automated.

    Intune customization profile scope
      By default only the Support information block of the default customization
      profile is written. The Branding block is opt-in via -ApplyIntuneVisualBranding.
      The Configuration block (enrollment availability, Company Portal app visibility,
      hidden remove/reset buttons, privacy messages) is deliberately left alone: those
      are per-client decisions, and the privacy statement URL in particular must not be
      copied from the Novix tenant. If you do want the Configuration block templated
      too, the property names need verifying against the beta API first - they are not
      covered by CIPP's standard, so they are not guessed at here.

    Required Microsoft Graph PowerShell modules
      Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Partner,
      Microsoft.Graph.Identity.Governance, Microsoft.Graph.Groups, Microsoft.Graph.Users

    Delegated access note
      Phase 2 signs in to the customer tenant with your partner credentials. The
      "Microsoft Graph Command Line Tools" enterprise application must be consented in
      the customer tenant on first use; the Application Administrator and Privileged
      Role Administrator roles in the GDAP role list below let you grant that consent.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$CustomerTenantId,

    [string[]]$GdapGroupMember = @(),

    [switch]$SkipGdap,
    [switch]$SkipBranding,
    [switch]$SkipSignInBranding,
    [switch]$SkipIntuneBranding,

    [switch]$ApplyIntuneVisualBranding,

    [string]$IntuneOrganizationName,

    [ValidateSet('P1Y', 'P2Y')]
    [string]$RelationshipDurationIso8601 = 'P2Y',

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

#region Novix support details -------------------------------------------------
# Single source of truth. Change a value here and it flows into the sign-in page
# text, the Intune profile, and the help desk checklist.

$Novix = [ordered]@{
    SupportName       = 'Novix IT Support'
    SupportPhone      = '+44 33 33 580 560'      # display form
    SupportPhoneDial  = '+443333580560'          # tel: link and Intune field
    SupportEmail      = 'support@novixit.co.uk'
    PortalUrl         = 'https://portal.novixit.co.uk'
    PortalLinkLabel   = 'Support Portal'         # link text on the sign-in page
    HelpDeskUrlLabel  = 'Novix IT Support Portal'
    IntuneSiteName    = 'Novix IT Client Portal'
    IntuneThemeColour = '#02102D'
    GroupPrefix       = 'M365 GDAP '
}

$GdapRoleDisplayNames = @(
    'Application Administrator'
    'Authentication Policy Administrator'
    'Billing Administrator'
    'Cloud App Security Administrator'
    'Cloud Device Administrator'
    'Domain Name Administrator'
    'Exchange Administrator'
    'Global Reader'
    'Helpdesk Administrator'
    'Intune Administrator'
    'Privileged Authentication Administrator'
    'Privileged Role Administrator'
    'Security Administrator'
    'SharePoint Administrator'
    'Teams Administrator'
    'User Administrator'
)

# Joined with LF - see the NOTES block above before changing this to a here-string.
$SignInPageTextLines = @(
    '**Forgotten Password or Username?**'
    ''
    'No worries! It happens to the best of us. Simply click on the "Forgotten my password" link on the sign-in page above. You will be prompted to enter the email address associated with your 365 account.'
    ''
    '**Trouble Signing In?**'
    ''
    "Ensure that you are using the correct password. Pay attention to upper and lower case characters, as our fields are case-sensitive. If you're still unable to sign in, try clearing your browser's cache and cookies or attempt signing in through a different browser."
    ''
    '**Browser Compatibility**'
    ''
    'For optimal user experience, ensure that your browser is up-to-date. 365 is compatible with the latest versions of Chrome, Firefox, Safari, and Edge.'
    ''
    '**Still, need assistance or have queries?**'
    ''
    ('Contact our support team through the [{0}]({1}), Call Us at [{2}](tel:{3}) or drop us an email at [{4}](mailto:{4}). We''re here to ensure your experience with 365 is seamless!' -f
        $Novix.PortalLinkLabel, $Novix.PortalUrl, $Novix.SupportPhone, $Novix.SupportPhoneDial, $Novix.SupportEmail)
)

$SignInPageTextMaxLength = 1024

#endregion

#region Helpers ---------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host ''
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Message)
    Write-Host "    [ok] $Message" -ForegroundColor Green
}

function Write-Skip {
    param([string]$Message)
    Write-Host "    [--] $Message" -ForegroundColor DarkGray
}

function Assert-RequiredModule {
    param([string[]]$Name)

    $missing = @($Name | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    if ($missing.Count -gt 0) {
        throw ("Missing required module(s): {0}. Install with: Install-Module {1} -Scope CurrentUser" -f
            ($missing -join ', '), ($missing -join ', '))
    }
}

function Convert-HexToRgbColor {
    <#
        Entra/Intune themeColor is an rgbColor object, not a hex string.
    #>
    param([Parameter(Mandatory)][string]$Hex)

    $clean = $Hex.TrimStart('#')
    if ($clean.Length -ne 6) { throw "Colour '$Hex' is not a 6-digit hex value." }

    return @{
        r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
        g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
        b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    }
}

function Invoke-GraphPatch {
    <#
        PATCH a resource. Graph rejects the whole payload if any single property name is
        unknown to the tenant's API version, so on failure retry property-by-property to
        apply what is valid and report precisely what was not.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Body,
        [string]$Label = 'resource'
    )

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body $Body -ErrorAction Stop | Out-Null
        return [pscustomobject]@{ Applied = @($Body.Keys); Failed = @{} }
    } catch {
        Write-Warning "Combined update of $Label failed: $($_.Exception.Message)"
        Write-Warning 'Retrying one property at a time to isolate the cause.'

        $applied = @()
        $failed = @{}
        foreach ($key in @($Body.Keys)) {
            try {
                Invoke-MgGraphRequest -Method PATCH -Uri $Uri -Body @{ $key = $Body[$key] } -ErrorAction Stop | Out-Null
                $applied += $key
            } catch {
                $failed[$key] = $_.Exception.Message
            }
        }
        return [pscustomobject]@{ Applied = $applied; Failed = $failed }
    }
}

function Set-ClipboardSafely {
    param([string]$Value)

    try {
        Set-Clipboard -Value $Value -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

#endregion

#region Phase 1 - partner tenant ---------------------------------------------

function Invoke-GdapOnboarding {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CustomerTenantId,
        [string[]]$GroupMember = @(),
        [string]$Duration = 'P2Y'
    )

    Write-Step 'Phase 1: connecting to the Novix partner tenant'

    Connect-MgGraph -Scopes @(
        'DelegatedAdminRelationship.ReadWrite.All'
        'Directory.ReadWrite.All'
        'Group.ReadWrite.All'
        'RoleManagement.Read.Directory'
        'User.Read.All'
    ) -NoWelcome

    $context = Get-MgContext
    Write-Ok "Signed in as $($context.Account) in tenant $($context.TenantId)"

    # --- Resolve role definitions --------------------------------------------
    Write-Step 'Resolving Entra role definitions'

    $roleDefinitions = @(
        Get-MgRoleManagementDirectoryRoleDefinition -All |
            Where-Object { $_.DisplayName -in $GdapRoleDisplayNames } |
            Select-Object DisplayName, Id
    )

    $missingRoles = @($GdapRoleDisplayNames | Where-Object { $_ -notin $roleDefinitions.DisplayName })
    if ($missingRoles.Count -gt 0) {
        throw "Role definition(s) not found in this tenant: $($missingRoles -join ', ')"
    }
    Write-Ok "$($roleDefinitions.Count) role definitions resolved"

    # --- Ensure one security group per role ----------------------------------
    Write-Step 'Ensuring GDAP security groups exist'

    $groupMap = @{}
    foreach ($role in $roleDefinitions) {
        $groupName = '{0}{1}' -f $Novix.GroupPrefix, $role.DisplayName
        $escaped = $groupName.Replace("'", "''")

        $group = Get-MgGroup -Filter "displayName eq '$escaped'" -Top 1 -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($group) {
            Write-Skip "$groupName (exists)"
        } elseif ($PSCmdlet.ShouldProcess($groupName, 'Create security group')) {
            $mailNickname = ($groupName -replace '[^a-zA-Z0-9]', '').ToLower()
            if ($mailNickname.Length -gt 60) { $mailNickname = $mailNickname.Substring(0, 60) }

            $group = New-MgGroup -DisplayName $groupName `
                -MailEnabled:$false `
                -MailNickname $mailNickname `
                -SecurityEnabled:$true
            Write-Ok "$groupName (created)"
        }

        if ($group) { $groupMap[$role.DisplayName] = $group }
    }

    if ($groupMap.Count -ne $roleDefinitions.Count) {
        Write-Warning 'Not every group is present (expected under -WhatIf). Stopping before relationship creation.'
        return
    }

    # --- Optional group membership -------------------------------------------
    if ($GroupMember.Count -gt 0) {
        Write-Step 'Adding members to GDAP groups'

        foreach ($upn in $GroupMember) {
            $user = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
            if (-not $user) {
                Write-Warning "User not found, skipping: $upn"
                continue
            }

            foreach ($roleName in $groupMap.Keys) {
                $group = $groupMap[$roleName]
                $already = Get-MgGroupMember -GroupId $group.Id -All |
                    Where-Object { $_.Id -eq $user.Id }

                if ($already) { continue }
                if (-not $PSCmdlet.ShouldProcess("$upn -> $($group.DisplayName)", 'Add group member')) { continue }

                New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)"
                } | Out-Null
            }
            Write-Ok "$upn added to all GDAP groups"
        }
    }

    # --- Reuse an existing relationship if there is one ----------------------
    Write-Step 'Checking for an existing GDAP relationship'

    $relationship = Get-MgTenantRelationshipDelegatedAdminRelationship -All |
        Where-Object {
            $_.Customer.TenantId -eq $CustomerTenantId -and
            $_.Status -in @('active', 'approvalPending', 'created')
        } |
        Sort-Object -Property Status |
        Select-Object -First 1

    if ($relationship) {
        Write-Ok "Reusing relationship '$($relationship.DisplayName)' (status: $($relationship.Status))"
    } else {
        $displayName = 'Novix-IT-Admin-Groups-{0}' -f $CustomerTenantId.Substring(0, 8)

        if (-not $PSCmdlet.ShouldProcess($displayName, 'Create GDAP relationship')) {
            Write-Warning 'Relationship not created (-WhatIf). Stopping phase 1.'
            return
        }

        $relationship = New-MgTenantRelationshipDelegatedAdminRelationship -BodyParameter @{
            displayName        = $displayName
            duration           = $Duration
            autoExtendDuration = 'P180D'
            customer           = @{ tenantId = $CustomerTenantId }
            accessDetails      = @{
                unifiedRoles = @($roleDefinitions | ForEach-Object { @{ roleDefinitionId = $_.Id } })
            }
        }
        Write-Ok "Created relationship $($relationship.Id)"
    }

    $relationshipId = $relationship.Id

    # --- Lock for approval ---------------------------------------------------
    if ($relationship.Status -eq 'created') {
        Write-Step 'Locking relationship for customer approval'

        New-MgTenantRelationshipDelegatedAdminRelationshipRequest `
            -DelegatedAdminRelationshipId $relationshipId `
            -BodyParameter @{ action = 'lockForApproval' } | Out-Null

        Write-Ok 'Locked for approval'
    }

    if ($relationship.Status -ne 'active') {
        $approvalUrl = "https://admin.microsoft.com/AdminPortal/Home#/partners/invitation/granularAdminRelationships/$relationshipId"

        Write-Host ''
        Write-Host 'Send this approval URL to the customer Global Administrator:' -ForegroundColor Yellow
        Write-Host "  $approvalUrl"

        if (Set-ClipboardSafely -Value $approvalUrl) {
            Write-Ok 'Approval URL copied to clipboard'
        } else {
            Write-Skip 'Clipboard unavailable on this host - copy the URL above manually'
        }

        # --- Wait for activation ---------------------------------------------
        Write-Step 'Waiting for the relationship to become active'

        $maxAttempts = 60
        $delaySeconds = 15
        $isActive = $false

        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            $current = Get-MgTenantRelationshipDelegatedAdminRelationship -DelegatedAdminRelationshipId $relationshipId
            Write-Host "    check $attempt/$maxAttempts : $($current.Status)"

            if ($current.Status -eq 'active') {
                $isActive = $true
                break
            }
            if ($current.Status -in @('terminated', 'terminationRequested')) {
                throw "Relationship entered status '$($current.Status)'. Investigate before re-running."
            }

            Start-Sleep -Seconds $delaySeconds
        }

        if (-not $isActive) {
            Write-Warning 'Relationship is not active yet. Once the customer approves it, re-run this script - it will pick up the existing relationship and continue.'
            return
        }
        Write-Ok 'Relationship is active'
    }

    # --- Access assignments --------------------------------------------------
    Write-Step 'Creating access assignments (one per role/group pair)'

    $existingAssignments = @(
        Get-MgTenantRelationshipDelegatedAdminRelationshipAccessAssignment `
            -DelegatedAdminRelationshipId $relationshipId -All -ErrorAction SilentlyContinue
    )

    $created = @()
    foreach ($role in $roleDefinitions) {
        $group = $groupMap[$role.DisplayName]

        $already = $existingAssignments | Where-Object {
            $_.AccessContainer.AccessContainerId -eq $group.Id
        }
        if ($already) {
            Write-Skip "$($role.DisplayName) (assignment exists)"
            continue
        }

        if (-not $PSCmdlet.ShouldProcess("$($group.DisplayName) -> $($role.DisplayName)", 'Create access assignment')) { continue }

        $assignment = New-MgTenantRelationshipDelegatedAdminRelationshipAccessAssignment `
            -DelegatedAdminRelationshipId $relationshipId `
            -BodyParameter @{
                accessContainer = @{
                    accessContainerId   = $group.Id
                    accessContainerType = 'securityGroup'
                }
                accessDetails   = @{
                    unifiedRoles = @(@{ roleDefinitionId = $role.Id })
                }
            }

        $created += [pscustomobject]@{
            Role   = $role.DisplayName
            Group  = $group.DisplayName
            Id     = $assignment.Id
            Status = $assignment.Status
        }
        Write-Ok "$($role.DisplayName)"
    }

    if ($created.Count -eq 0) {
        Write-Ok 'All access assignments were already in place'
        return
    }

    # --- Wait for assignments to activate ------------------------------------
    Write-Step 'Waiting for access assignments to activate'

    for ($attempt = 1; $attempt -le 40; $attempt++) {
        $statuses = @(
            Get-MgTenantRelationshipDelegatedAdminRelationshipAccessAssignment `
                -DelegatedAdminRelationshipId $relationshipId -All
        ).Status

        $pending = @($statuses | Where-Object { $_ -ne 'active' })
        Write-Host "    check $attempt/40 : $($statuses.Count) assignment(s), $($pending.Count) not yet active"

        if ($pending.Count -eq 0) {
            Write-Ok 'All access assignments are active'
            return
        }
        Start-Sleep -Seconds 10
    }

    Write-Warning 'Assignments were created but are not all active yet. They usually settle within a few minutes.'
}

#endregion

#region Phase 2 - customer tenant --------------------------------------------

function Set-SignInPageBranding {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$Text
    )

    Write-Step 'Setting Entra ID sign-in page text'

    if ($Text.Length -gt $SignInPageTextMaxLength) {
        throw ("Sign-in page text is {0} characters; Entra ID allows {1}. Trim the wording in `$SignInPageTextLines." -f
            $Text.Length, $SignInPageTextMaxLength)
    }
    Write-Ok "Text length $($Text.Length)/$SignInPageTextMaxLength characters"

    if (-not $PSCmdlet.ShouldProcess($TenantId, 'Set sign-in page text')) { return }

    $brandingUri = "/v1.0/organization/$TenantId/branding"
    $body = @{ signInPageText = $Text }

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $brandingUri -Body $body -ErrorAction Stop | Out-Null
        Write-Ok 'Updated the default branding object'
    } catch {
        # A tenant that has never had branding configured has no default branding
        # object to PATCH. Create it as the default ("0") localization instead.
        Write-Skip "PATCH failed ($($_.Exception.Message.Trim()))"
        Write-Skip 'No default branding object yet - creating it'

        Invoke-MgGraphRequest -Method POST -Uri "$brandingUri/localizations" `
            -Body ($body + @{ id = '0' }) -ErrorAction Stop | Out-Null
        Write-Ok 'Created the default branding object'
    }

    # Read it back so we know it actually landed rather than trusting a 204.
    try {
        $stored = Invoke-MgGraphRequest -Method GET -Uri $brandingUri -ErrorAction Stop
        if ($stored.signInPageText -eq $Text) {
            Write-Ok 'Verified: stored text matches'
        } else {
            Write-Warning "Stored text does not match what was sent. Stored length: $($stored.signInPageText.Length)"
        }
    } catch {
        Write-Warning "Could not read branding back for verification: $($_.Exception.Message)"
    }
}

function Set-IntuneSupportBranding {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$OrganizationName,
        [switch]$ApplyVisualBranding
    )

    Write-Step 'Setting the Intune Company Portal branding profile'

    $profiles = $null
    try {
        $profiles = Invoke-MgGraphRequest -Method GET -Uri '/beta/deviceManagement/intuneBrandingProfiles' -ErrorAction Stop
    } catch {
        Write-Warning "Could not read Intune branding profiles: $($_.Exception.Message)"
        Write-Warning 'If Intune has never been opened in this tenant there is nothing to patch yet. Open the Intune admin center once, then re-run with -SkipGdap.'
        return
    }

    $default = @($profiles.value) | Where-Object { $_.isDefaultProfile } | Select-Object -First 1
    if (-not $default) {
        Write-Warning 'No default Intune branding profile found. Skipping.'
        return
    }
    Write-Ok "Default profile: $($default.profileName) ($($default.id))"

    if (-not $PSCmdlet.ShouldProcess($default.profileName, 'Set support information')) { return }

    $uri = "/beta/deviceManagement/intuneBrandingProfiles/$($default.id)"

    # Property names confirmed against CIPP's own intuneBrandingProfile standard
    # (src/data/standards.json -> standards.intuneBrandingProfile).
    #
    # privacyUrl is deliberately absent: the Novix privacy statement does not apply to
    # client tenants, and we must not overwrite a value the client has set themselves.
    $support = @{
        contactITName         = $Novix.SupportName
        contactITPhoneNumber  = $Novix.SupportPhoneDial
        contactITEmailAddress = $Novix.SupportEmail
        onlineSupportSiteName = $Novix.IntuneSiteName
        onlineSupportSiteUrl  = $Novix.PortalUrl
    }

    $result = Invoke-GraphPatch -Uri $uri -Body $support -Label 'Intune support information'
    foreach ($key in $result.Applied) { Write-Ok $key }
    foreach ($key in $result.Failed.Keys) { Write-Warning "$key : $($result.Failed[$key])" }

    if (-not $ApplyVisualBranding) {
        Write-Skip 'Visual branding (org name, theme colour, header) left as-is - pass -ApplyIntuneVisualBranding to set it'
        return
    }

    # Opt-in: this overwrites the client's own Company Portal identity, including
    # hiding their logo in favour of the organization name. themeColor is patched on
    # its own so a rejection there cannot take the rest down with it.
    $visual = @{
        displayName               = $OrganizationName
        showLogo                  = $false
        showDisplayNameNextToLogo = $true
    }

    $visualResult = Invoke-GraphPatch -Uri $uri -Body $visual -Label 'Intune visual branding'
    foreach ($key in $visualResult.Applied) { Write-Ok $key }
    foreach ($key in $visualResult.Failed.Keys) { Write-Warning "$key : $($visualResult.Failed[$key])" }

    $theme = @{ themeColor = Convert-HexToRgbColor -Hex $Novix.IntuneThemeColour }
    $themeResult = Invoke-GraphPatch -Uri $uri -Body $theme -Label 'Intune theme colour'
    foreach ($key in $themeResult.Applied) { Write-Ok "$key ($($Novix.IntuneThemeColour))" }
    foreach ($key in $themeResult.Failed.Keys) { Write-Warning "$key : $($themeResult.Failed[$key])" }
}

function Show-HelpDeskInformation {
    param([string]$DefaultDomain)

    Write-Step 'Microsoft 365 help desk information (manual step)'

    Write-Host '    No supported API exists for this form - see NOTES in this script.' -ForegroundColor DarkGray
    Write-Host '    Microsoft 365 admin center > Settings > Org settings >' -ForegroundColor DarkGray
    Write-Host '    Organization profile > Help desk information' -ForegroundColor DarkGray
    Write-Host ''

    $values = [ordered]@{
        'Add your help desk contact information' = 'Checked'
        'Title'                                  = $Novix.SupportName
        'Phone'                                  = $Novix.SupportPhoneDial
        'Email'                                  = $Novix.SupportEmail
        'URL'                                    = $Novix.PortalUrl
        'URL label'                              = $Novix.HelpDeskUrlLabel
    }

    foreach ($key in $values.Keys) {
        Write-Host ('    {0,-40} {1}' -f $key, $values[$key])
    }

    $adminUrl = 'https://admin.microsoft.com/Adminportal/Home#/Settings/OrgSettings'
    if ($DefaultDomain) {
        $adminUrl = "https://admin.microsoft.com/?delegatedOrg=$DefaultDomain#/Settings/OrgSettings"
    }

    Write-Host ''
    Write-Host "    $adminUrl" -ForegroundColor Yellow
}

function Invoke-CustomerTenantBranding {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CustomerTenantId,
        [string]$OrganizationNameOverride,
        [switch]$SkipSignIn,
        [switch]$SkipIntune,
        [switch]$ApplyIntuneVisual,
        [switch]$NoConfirm
    )

    Write-Step "Phase 2: connecting to customer tenant $CustomerTenantId (delegated)"

    Connect-MgGraph -TenantId $CustomerTenantId -Scopes @(
        'Organization.ReadWrite.All'
        'OrganizationalBranding.ReadWrite.All'
        'DeviceManagementApps.ReadWrite.All'
    ) -NoWelcome

    $context = Get-MgContext
    if ($context.TenantId -ne $CustomerTenantId) {
        throw "Connected to tenant $($context.TenantId) but expected $CustomerTenantId. Aborting before writing anything."
    }

    $org = (Invoke-MgGraphRequest -Method GET -Uri '/v1.0/organization').value | Select-Object -First 1
    $orgName = $org.displayName
    $defaultDomain = ($org.verifiedDomains | Where-Object { $_.isDefault }).name

    Write-Ok "Tenant: $orgName ($defaultDomain)"
    Write-Ok "Signed in as $($context.Account)"

    if (-not $NoConfirm -and -not $WhatIfPreference) {
        Write-Host ''
        $answer = Read-Host "About to apply Novix branding to '$orgName'. Continue? (y/N)"
        if ($answer -notmatch '^(y|yes)$') {
            Write-Warning 'Aborted by operator. Nothing was changed.'
            return
        }
    }

    if ($SkipSignIn) {
        Write-Skip 'Sign-in page text skipped (-SkipSignInBranding)'
    } else {
        Set-SignInPageBranding -TenantId $CustomerTenantId -Text ($SignInPageTextLines -join "`n")
    }

    if ($SkipIntune) {
        Write-Skip 'Intune branding profile skipped (-SkipIntuneBranding)'
    } else {
        $intuneOrgName = if ($OrganizationNameOverride) { $OrganizationNameOverride } else { $orgName }
        Set-IntuneSupportBranding -OrganizationName $intuneOrgName -ApplyVisualBranding:$ApplyIntuneVisual
    }

    Show-HelpDeskInformation -DefaultDomain $defaultDomain
}

#endregion

#region Main ------------------------------------------------------------------

Assert-RequiredModule -Name @('Microsoft.Graph.Authentication')

if (-not $SkipGdap) {
    Assert-RequiredModule -Name @(
        'Microsoft.Graph.Identity.Partner'
        'Microsoft.Graph.Identity.Governance'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Users'
    )

    Import-Module Microsoft.Graph.Identity.Partner, Microsoft.Graph.Identity.Governance,
        Microsoft.Graph.Groups, Microsoft.Graph.Users

    Invoke-GdapOnboarding -CustomerTenantId $CustomerTenantId `
        -GroupMember $GdapGroupMember `
        -Duration $RelationshipDurationIso8601
} else {
    Write-Skip 'Phase 1 skipped (-SkipGdap)'
}

if ($SkipBranding) {
    Write-Skip 'Phase 2 skipped (-SkipBranding)'
} else {
    Invoke-CustomerTenantBranding -CustomerTenantId $CustomerTenantId `
        -OrganizationNameOverride $IntuneOrganizationName `
        -SkipSignIn:$SkipSignInBranding `
        -SkipIntune:$SkipIntuneBranding `
        -ApplyIntuneVisual:$ApplyIntuneVisualBranding `
        -NoConfirm:$Force
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green

#endregion
