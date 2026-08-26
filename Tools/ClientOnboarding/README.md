# Client tenant onboarding

`Invoke-NovixClientOnboarding.ps1` onboards a client tenant in two phases and is safe to
re-run — every write is idempotent, and `-WhatIf` previews the whole thing without
touching anything.

## What it does

**Phase 1 — Novix partner tenant**

- Ensures one security group per GDAP role exists (`M365 GDAP <Role>`), for the 16 roles
  listed in the script.
- Optionally adds named users to every group (`-GdapGroupMember`).
- Creates the delegated admin (GDAP) relationship, locks it for approval, prints and
  copies the approval URL, then waits for the customer to accept.
- Creates one access assignment per role/group pair and waits for them to activate.

**Phase 2 — client tenant (delegated, over the GDAP relationship)**

| Setting | Where it lives | Automated |
| --- | --- | --- |
| Sign-in page text | Entra ID → Company branding | Yes |
| Company Portal support information | Intune → Tenant admin → Customization | Yes |
| Company Portal visual branding | Intune → Tenant admin → Customization | Opt-in, `-ApplyIntuneVisualBranding` |
| Help desk information | M365 admin → Org settings → Organization profile | No — printed for paste, see below |

## Usage

```powershell
# Full onboarding
.\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '00000000-1111-2222-3333-444444444444'

# Preview everything, change nothing
.\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -WhatIf

# Branding only, for a client already onboarded
.\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -SkipGdap

# GDAP only
.\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -SkipBranding

# Also push the Novix Company Portal look (overwrites the client's own)
.\Invoke-NovixClientOnboarding.ps1 -CustomerTenantId '<guid>' -SkipGdap -ApplyIntuneVisualBranding
```

If the customer has not approved the relationship by the time the wait loop gives up,
just re-run the script once they have — it picks up the existing relationship and carries
on from the access assignments.

## Prerequisites

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Identity.Partner,
    Microsoft.Graph.Identity.Governance, Microsoft.Graph.Groups, Microsoft.Graph.Users -Scope CurrentUser
```

You sign in with **your own Novix account** for both phases — never a client account.
Phase 2 reaches the client tenant through the GDAP relationship, not through client
credentials. On first use the **Microsoft Graph Command Line Tools** enterprise
application needs consent in the client tenant; the Application Administrator and
Privileged Role Administrator roles in the GDAP role list cover granting it.

### Your account needs Admin agent in Partner Center

Creating a GDAP relationship requires the **Admin agent** role in Partner Center.
Global Administrator in the Novix tenant is *not* sufficient — and nothing else in
phase 1 needs it, so without the role the script reads roles and creates all 16 groups
quite happily before failing on one bare error:

```
New-MgTenantRelationshipDelegatedAdminRelationship : Access to the resource is restricted.
Status: 403 (Forbidden)   ErrorCode: forbidden
```

That is what the pre-flight exists to prevent. It checks, in order of how conclusive
each check is:

1. **Granted scopes** — read off the token, not the list you asked for. Requesting a
   scope and being granted it are different things. Hard failure.
2. **A read probe** against delegated admin relationships. This exercises the same
   Partner Center authorisation as creating one, so a 403 here identifies the problem
   before a single group is created. Hard failure, with the fix in the message.
3. **`AdminAgents` group membership** — a heuristic, because nested membership and
   renamed groups aren't detected. Warns and asks rather than failing.

Assign the role under Partner Center → Settings → Account settings → User management,
then `Disconnect-MgGraph` and sign in again — role membership is cached in the token.

### Watch which account WAM picks

On Windows, `Connect-MgGraph` uses Web Account Manager by default and will often sign
you in silently as whatever account Windows is already using, with no prompt. If you
hold both a standard and an admin Novix account, that can quietly be the wrong one.
The pre-flight prints the account it ended up as. To choose explicitly:

```powershell
Disconnect-MgGraph
Connect-MgGraph -Scopes "DelegatedAdminRelationship.ReadWrite.All","Directory.ReadWrite.All",`
    "Group.ReadWrite.All","RoleManagement.Read.Directory","User.Read.All" -UseDeviceCode
```

Device code flow bypasses WAM entirely. (Older module versions call that switch
`-UseDeviceAuthentication`.)

## Changing the Novix support details

Everything lives in the `$Novix` block at the top of the script — phone, email, portal
URL, support name, theme colour. Change a value there and it flows into the sign-in page
text, the Intune profile, and the help desk checklist.

## Two things to know before editing the sign-in page text

1. **It is 1022 characters against a 1024 cap.** Two characters of headroom. The script
   hard-fails with the exact count if an edit pushes it over, rather than letting Graph
   reject it with a vague validation error.
2. **It is built from an array joined with `` "`n" ``, not a here-string.** A here-string
   picks up the file's line endings; with CRLF the same text measures 1036 characters and
   the API call fails. Keep the array-join.

Entra renders a limited markdown subset in this field — bold, italics, and
`[label](url)` links. That is why the text can carry formatted headings and a working
support portal link.

## Why help desk information is not automated

Microsoft exposes no supported API for **Settings → Org settings → Organization profile →
Help desk information**. CIPP has no standard for it either, for the same reason. The
only programmatic route is the undocumented `admin.microsoft.com` portal API, which needs
a token for a non-Graph audience and can change without notice.

Rather than ship a guessed endpoint, phase 2 prints the exact values and a delegated
admin link, so it is a short paste. To automate it properly, capture the real request
from browser dev tools while saving that form once, and it can be wired in.

## Property name provenance

The Intune branding properties (`contactITName`, `contactITPhoneNumber`,
`contactITEmailAddress`, `onlineSupportSiteName`, `onlineSupportSiteUrl`, `displayName`,
`showLogo`, `showDisplayNameNextToLogo`, `privacyUrl`) are the ones CIPP itself uses in
its `standards.intuneBrandingProfile` standard — see `src/data/standards.json`.

The Configuration block of the customization profile (enrollment availability, Company
Portal app visibility, hidden remove/reset buttons, privacy messages) is **not** set.
Those are per-client decisions, the Novix privacy statement URL must not be copied into
client tenants, and their beta property names are not covered by CIPP's standard, so they
are not guessed at here.

Where a property name might not be accepted by a given tenant's API version, the script
retries the patch one property at a time and reports exactly which ones were rejected,
rather than losing the whole payload to a single bad field.
