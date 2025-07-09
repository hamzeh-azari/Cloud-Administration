# Microsoft 365 Email Activity Reporter to SQL Server 📊

This PowerShell project retrieves email activity details of selected users from **Microsoft 365** via the **Microsoft Graph API**, and inserts them into a **SQL Server** table — making it possible to visualize and analyze the data in **Power BI** or any reporting system.

---

## 🔧 Use Cases

- Track email send/receive counts for specific users or departments
- Monitor internal communication trends
- Feed clean data into **Power BI dashboards** for analysis
- Detect anomalies or identify inactive users

---

## 💡 What This Script Does

1. Authenticates to Microsoft Graph using **client credentials (OAuth2)**  
2. Downloads **email activity user details** for the past 7 days  
3. Filters the report for specific users (defined by email)  
4. Automatically creates the required table in SQL Server if it doesn't exist  
5. Inserts the filtered data (username, send/receive counts, date range) into SQL  
6. Prepares clean, structured data ready for **Power BI visualization**

---

## 📌 Requirements

- PowerShell 5.1+
- `ImportExcel` PowerShell module  
- Access to SQL Server (on-prem or cloud)
- Azure AD App with the following permission:
  - `Reports.Read.All` (Application permission in Microsoft Graph API)
- Admin consent granted in Azure portal

---

## 🔐 Configuration Explained

You’ll need to replace the following placeholders in the script:

| Variable          | What it is | Where to get it |
|------------------|------------|-----------------|
| `$tenantId`      | Azure AD Tenant ID (GUID) | Azure Portal → Azure AD → Overview |
| `$clientId`      | Azure App’s Client ID | Azure Portal → App registrations |
| `$clientSecret`  | Secret string for the Azure App | Certificates & secrets section of the App |
| `$emailsToKeep`  | List of user emails to track | Your choice (team members, department, etc.) |
| `$sqlServer`     | SQL Server instance name | e.g. `localhost`, `SERVER\INSTANCE` |
| `$sqlDatabase`   | Name of your database | e.g. `ReportsDB` |
| `$sqlUser` / `$sqlPassword` | Login credentials | SQL Authentication (or adapt to Windows Auth) |

---

## 🧭 Setup Steps (One-Time)

### 1. Register an App in Azure
- Go to **Azure Portal > Azure Active Directory > App registrations**
- Click **New registration**
- Give it a name (e.g. `M365 Reporter App`)
- Choose *Single tenant*
- After creation, copy:
  - **Application (client) ID →** `$clientId`
  - **Directory (tenant) ID →** `$tenantId`

### 2. Create a Client Secret
- In the same App → **Certificates & secrets**
- Click **New client secret**
- Choose expiration (6 or 12 months)
- Copy the secret value → `$clientSecret`

### 3. Add Microsoft Graph API Permission
- In your App → **API permissions > Add permission**
- Choose **Microsoft Graph > Application permissions**
- Search for and select: `Reports.Read.All`
- Click **Grant admin consent**

---

## 🚀 How to Run the Script

1. Open PowerShell as Administrator  
2. Make sure the `ImportExcel` module is available:
   ```powershell
   Install-Module -Name ImportExcel -Scope CurrentUser -Force
