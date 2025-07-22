<#
.SYNOPSIS
    Exports Microsoft 365 MFA methods per user categorized by method type into an Excel report.

.DESCRIPTION
    This PowerShell script connects to Microsoft Graph API and retrieves each Azure AD user's registered multi-factor authentication (MFA) methods.
    The results are categorized into standard types such as Password, Microsoft Authenticator, Phone (Call/SMS), Software OATH (e.g., Google Authenticator), and any others.

    ✅ Use this script to:
        - Audit MFA usage across your organization
        - Identify users without strong MFA configured
        - Filter and sort MFA types easily in Excel

    The output is exported to a structured Excel file for further reporting or security review.

.NOTES
    Author: Hamzeh Azari Hashjin
    Date: 2025-07-22
    Requirements:
        - Microsoft Graph PowerShell SDK
        - ImportExcel module
        - Permissions: User.Read.All, UserAuthenticationMethod.Read.All
        - PowerShell 5.1+
#>

# --- Module Installation ---
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force

# --- Connect to Graph ---
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All"

# --- Retrieve Users ---
$users = Get-MgUser -All
$results = @()

foreach ($user in $users) {
    Write-Host "User: $($user.UserPrincipalName)" -ForegroundColor Cyan
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $user.Id
        $methodTypes = $methods | ForEach-Object {
            $_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', ''
        }

        # Initialize friendly name slots
        $password = ""
        $authenticator = ""
        $phone = ""
        $softwareOath = ""
        $others = @()

        foreach ($type in $methodTypes) {
            switch ($type) {
                "passwordAuthenticationMethod" { $password = "Password" }
                "microsoftAuthenticatorAuthenticationMethod" { $authenticator = "Microsoft Authenticator" }
                "phoneAuthenticationMethod" { $phone = "Phone (Call/SMS)" }
                "softwareOathAuthenticationMethod" { $softwareOath = "OTP Authenticator" }
                default { $others += $type }
            }
            Write-Host "  MFA Method: $type"
        }

        $obj = [ordered]@{
            DisplayName        = $user.DisplayName
            UserPrincipalName  = $user.UserPrincipalName
            Password           = $password
            Authenticator      = $authenticator
            Phone              = $phone
            SoftwareOath       = $softwareOath
        }

        for ($i = 0; $i -lt $others.Count; $i++) {
            $obj["Other$($i+1)"] = $others[$i]
        }

        $results += New-Object PSObject -Property $obj
    } catch {
        Write-Host "  Unable to retrieve MFA methods" -ForegroundColor Red
        $results += [PSCustomObject]@{
            DisplayName        = $user.DisplayName
            UserPrincipalName  = $user.UserPrincipalName
            Password           = ""
            Authenticator      = ""
            Phone              = ""
            SoftwareOath       = ""
            Other1             = "Error or No Access"
        }
    }
}

# --- Export to Excel ---
$exportPath = "$env:USERPROFILE\Desktop\MFA_Methods_By_Type.xlsx"
$results | Export-Excel -Path $exportPath -AutoSize -Title "MFA Methods By Type"

Write-Host "\nDone. Report exported to: $exportPath" -ForegroundColor Green
