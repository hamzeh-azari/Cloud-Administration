<#
.SYNOPSIS
    Retrieves Microsoft 365 email activity reports via Microsoft Graph API, filters specific users, and inserts the data into a SQL Server table.

.DESCRIPTION
    This script authenticates using client credentials, downloads the email activity report for the past 7 days,
    filters it for a predefined list of users, and inserts the relevant data into a SQL Server database.

    ✅ You can use the resulting data to:
        - Monitor user activity (email send/receive counts)
        - Visualize the activity trends in Power BI dashboards
        - Analyze user behavior and communication patterns over time

    This is especially useful for organizations that want to track productivity, identify inactive users, or detect anomalies.

.NOTES
    Author: Hamzeh Azari Hashjin
    Date: 2025-07-09
    Requirements:
        - Azure AD App with Microsoft Graph API permissions (Reports.Read.All)
        - PowerShell 5.1+
        - ImportExcel module
        - SQL Server access
#>

# ============================
# 1. Ensure ImportExcel module is available (used for handling CSV files)
# ============================
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# ============================
# 2. Configuration
# ============================

# Replace with your Azure AD tenant ID (GUID format)
$tenantId     = '<TENANT_ID>'

# Replace with your Azure AD application's client ID (from App registrations)
$clientId     = '<CLIENT_ID>'

# Replace with your Azure AD application's client secret
$clientSecret = '<CLIENT_SECRET>'

# List of user email addresses to filter from the report
$emailsToKeep = @(
    "email1@example.com",  # Replace with real user email
    "email2@example.com",
    "email3@example.com"
)

# Specify the reporting period (D7 = past 7 days, D30 = past 30 days)
$reportPeriod = 'D7'

# File path to save the raw CSV report
$rawReportPath = "$PWD\EmailActivity_Raw.csv"

# SQL Server connection details
$sqlServer   = "<SQL_SERVER>"        # e.g. "localhost", "SERVERNAME\INSTANCE"
$sqlDatabase = "<DATABASE_NAME>"     # e.g. "ReportsDB"
$sqlUser     = "<USERNAME>"          # SQL login user
$sqlPassword = "<PASSWORD>"          # SQL login password
$tableName   = "EmailActivity"       # Name of the destination table

# ============================
# 3. Acquire Microsoft Graph API Access Token
# ============================
$tokenBody = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody
$accessToken = $tokenResponse.access_token

# ============================
# 4. Download Email Activity Report
# ============================
$headers = @{
    Authorization = "Bearer $accessToken"
    Accept        = "application/octet-stream"
}

$reportUri = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='$reportPeriod')"

Invoke-RestMethod -Headers $headers -Uri $reportUri -Method Get -OutFile $rawReportPath

# ============================
# 5. Filter the CSV Report
# ============================
$csv = Import-Csv $rawReportPath

# Normalize and filter only the users in $emailsToKeep
$filtered = $csv | Where-Object {
    $email = $_."User Principal Name".Trim().ToLower()
    $emailsToKeep -contains $email
}

# ============================
# 6. Connect to SQL Server
# ============================
$connectionString = "Server=$sqlServer;Database=$sqlDatabase;User Id=$sqlUser;Password=$sqlPassword;"
$connection = New-Object System.Data.SqlClient.SqlConnection $connectionString
$connection.Open()

# ============================
# 7. Ensure the Target Table Exists
# ============================
$tableCheck = @"
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$tableName')
BEGIN
    CREATE TABLE [$tableName] (
        [User] NVARCHAR(255),
        [SendCount] INT,
        [ReceiveCount] INT,
        [StartDate] DATE,
        [EndDate] DATE
    );
END
"@

$command = $connection.CreateCommand()
$command.CommandText = $tableCheck
$command.ExecuteNonQuery()

# ============================
# 8. Insert Filtered Data into SQL Table
# ============================
$startDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd")
$endDate   = (Get-Date).ToString("yyyy-MM-dd")

foreach ($row in $filtered) {
    $user         = $row."User Principal Name".Replace("'", "''")  # Escape single quotes
    $sendCount    = [int]$row."Send Count"
    $receiveCount = [int]$row."Receive Count"

    $insertQuery = @"
INSERT INTO [$tableName] ([User], [SendCount], [ReceiveCount], [StartDate], [EndDate])
VALUES (N'$user', $sendCount, $receiveCount, '$startDate', '$endDate');
"@

    $command.CommandText = $insertQuery
    $command.ExecuteNonQuery()
}

# ============================
# 9. Cleanup and Final Message
# ============================
$connection.Close()
Write-Output "✅ Email activity data successfully inserted into [$tableName]."
