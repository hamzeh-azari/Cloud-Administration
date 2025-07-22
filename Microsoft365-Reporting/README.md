# Export Microsoft 365 Email Activity to SQL Server

This PowerShell project retrieves email activity details of selected users from Microsoft 365 via the Microsoft Graph API, and inserts them into a SQL Server table — making it possible to visualize and analyze the data in Power BI or any reporting system.

---

## 💡 Features
- Connects securely to Microsoft Graph using client credentials (OAuth2)
- Downloads Microsoft 365 email activity user detail report for the past 7 days
- Filters the report for specific users based on a predefined list
- Automatically creates the destination table in SQL Server if it doesn't exist
- Inserts relevant data (user, send/receive counts, date range) into SQL
- Generates clean, structured data ready for Power BI or Excel

---

## 🔧 Use Cases
 
- Track email send/receive counts for specific users or departments
- Monitor internal communication trends
- Feed clean data into **Power BI dashboards** for analysis
- Detect anomalies or identify inactive users

---

## 📦 Requirements
- PowerShell 5.1 or later
- ImportExcel PowerShell module
- Access to SQL Server (on-prem or cloud)
- Azure AD App with the following permission:
  - `Reports.Read.All` (Application permission in Microsoft Graph)
  - Admin consent granted in Azure portal

## ⚙️ Configuration
You'll need to replace the following variables inside the script:

| Variable         | What it is                                | Where to find it |
|------------------|--------------------------------------------|------------------|
| `$tenantId`      | Azure AD Tenant ID (GUID)                 | Azure Portal → Azure AD → Overview |
| `$clientId`      | Azure App’s Client ID                     | Azure Portal → App registrations |
| `$clientSecret`  | Client secret from Azure App              | Certificates & secrets section |
| `$emailsToKeep`  | List of user emails to track              | Your own list (team, dept, etc.) |
| `$sqlServer`     | SQL Server instance name                  | e.g. `localhost`, `SERVER\INSTANCE` |
| `$sqlDatabase`   | SQL Database name                         | e.g. `ReportsDB` |
| `$sqlUser` / `$sqlPassword` | SQL login credentials         | SQL Authentication info |

## 🧭 One-Time Setup

### 1. Register an App in Azure
- Go to Azure Portal > Azure AD > App registrations > New registration
- Name it (e.g. `M365 Reporter App`), choose *Single tenant*
- After creation, copy:
  - Application (client) ID → `$clientId`
  - Directory (tenant) ID → `$tenantId`

### 2. Create a Client Secret
- In the App → **Certificates & secrets** > New client secret
- Choose an expiration period (6 or 12 months)
- Copy and store the secret value → `$clientSecret`

### 3. Add Microsoft Graph API Permissions
- App > API Permissions > Add permission > Microsoft Graph
- Choose **Application permissions** > Search `Reports.Read.All`
- Add & **Grant admin consent**

## 🚀 How to Run the Script
1. Open PowerShell as Administrator
2. Ensure ImportExcel module is installed:
```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```
3. Modify the script with your credentials & users
4. Run it:
```powershell
.\EmailActivity-To-SQL.ps1
```

After execution, your data will be inserted into the SQL Server table (e.g., `EmailActivity`).

## 📈 Power BI Integration
Once your SQL table is populated:
- Open **Power BI Desktop**
- Click **Home > Get Data > SQL Server**
- Connect to your server and select the table
- Create charts such as:
  - Send/Receive counts per user
  - Weekly trends
  - Communication ranking and comparisons

## 📁 Project Structure
```
m365-email-to-sql/
├── EmailActivity-To-SQL.ps1
└── README.md
```

## 🛡️ Warning
⚠️ This script writes data directly to your SQL Server.  
**Always test in a development environment before using in production.**

## 🤝 Contributing
Pull requests, feature ideas or bug reports are welcome — especially if you add support for:
- Windows Authentication to SQL Server
- Custom time ranges (e.g., D30, D90)
- Exporting the data to Excel or SharePoint

## 📌 Author
**Hamzeh Azari Hashjin**  
☁️ Cloud & Systems Admin | 💻 12+ years in Hosting & Infrastructure  
📍 Based in Montreal, Canada  
🌐 [LinkedIn Profile](https://www.linkedin.com/in/hamzeh-azari/)

## 🛡️ License
MIT License. Feel free to use, adapt and share with credit.

