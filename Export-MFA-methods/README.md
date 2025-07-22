# Export MFA Methods by Type

This PowerShell script connects to Microsoft Graph API and exports all Azure AD users' registered MFA (Multi-Factor Authentication) methods into an Excel report. The output categorizes authentication methods by type (e.g., Password, Microsoft Authenticator, Phone, etc.) to allow easy filtering and reporting.

---

## 🔧 Features

* Connects securely to Microsoft Graph using delegated permissions
* Categorizes each user's MFA methods into separate columns:

  * Password
  * Microsoft Authenticator App
  * Phone (Call/SMS)
  * Software OATH (e.g., Google Authenticator)
  * Others (FIDO2, TemporaryAccessPass, etc.)
* Friendly method names for easier interpretation
* Exports to Excel (`MFA_Methods_By_Type.xlsx`) with sortable/filterable structure

## 📦 Requirements

* PowerShell 5.1 or later
* Microsoft.Graph PowerShell SDK
* ImportExcel PowerShell module
* Admin account with appropriate Graph API permissions

## 🛠️ Usage

Install required modules:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

Run the script:

```powershell
.\Export-MFA-Methods.ps1
```

After execution, the Excel file will be created on your Desktop.

## 🔐 Microsoft Graph Permissions

The following delegated permissions are required and must be admin-consented:

* `User.Read.All`
* `UserAuthenticationMethod.Read.All`

## 📁 Project Structure

```
export-mfa-methods/
├── Export-MFA-Methods.ps1
└── README.md
```

## 🛡️ Warning

⚠️ This script reads sensitive user authentication method data from Microsoft Graph.
**Always ensure access is restricted and tested in a secure environment before production use.**

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## 📌 Author

**Hamzeh Azari Hashjin**
☁️ Cloud & Systems Admin | 💻 12+ years in Hosting & Infrastructure
📍 Based in Montreal, Canada
🌐 [LinkedIn Profile](https://www.linkedin.com/in/hamzeh-azari/)

## 🛡️ License

MIT License. Feel free to use, adapt and share with credit.
