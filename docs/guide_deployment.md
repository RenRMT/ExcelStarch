# Deploying the Excel Add-in Safely Within an Organization

This guide covers best practices for safely deploying the Excel chart add-in across your organization, including digital signing, Trust Center configuration, deployment methods, and access control.

---

## Overview

Excel add-ins run with the same permissions as the user and can interact with worksheet data. Safe deployment requires:
1. **Digital signing** to establish trust and prevent tampering
2. **Trust Center configuration** to permit the add-in to run
3. **Deployment strategy** suited to your IT infrastructure
4. **Access control** to limit usage to authorized users

This guide addresses all four.

---

## Part 1: Digital Signing

### Why sign the add-in?

- **Authenticity:** Users can verify the add-in came from your organization
- **Integrity:** A signature prevents tampering — if the file is modified, the signature breaks
- **Trust Center:** Windows will trust a properly signed add-in; unsigned add-ins often trigger security warnings
- **Revocation:** If a signed add-in is compromised, you can revoke its certificate

### What you need

1. A **code-signing certificate** (.pfx file) issued by your organization or a trusted certificate authority
   - If your organization has an internal PKI (Public Key Infrastructure), request a code-signing certificate from your IT/Security team
   - Alternatively, purchase one from a public CA (DigiCert, Sectigo, etc.) — this allows users on any machine to verify the signature without special setup
2. The **certificate password** (if the certificate is password-protected)

### Step 1: Obtain or create a certificate

#### Option A: Use your organization's internal certificate (recommended for internal-only distribution)

Contact your IT/Security team and request a **code-signing certificate** in `.pfx` format. This certificate should:
- Be issued by your organization's Certificate Authority (CA)
- Have the **Code Signing** Enhanced Key Usage (EKU)
- Be valid for at least the intended deployment period

Store the `.pfx` file securely (e.g., on an encrypted drive or in a secure key vault). You will need its password to sign files.

#### Option B: Purchase a public code-signing certificate (recommended for external or long-term distribution)

If your add-in may be distributed outside your organization, purchase a code-signing certificate from a trusted CA:
- [DigiCert Code Signing Certificates](https://www.digicert.com/code-signing/code-signing-certificates)
- [Sectigo Code Signing Certificates](https://sectigo.com/ssl-certificates-tls/code-signing)
- [GlobalSign Code Signing](https://www.globalsign.com/en/code-signing)

Once purchased and installed on your machine, the certificate will appear in the system certificate store.

### Step 2: Sign the .xlam file

You can sign the `.xlam` file using **PowerShell** (built-in to Windows).

#### Sign using PowerShell

1. Open PowerShell as Administrator.
2. Navigate to the folder containing the `.xlam` file:
   ```powershell
   cd "C:\path\to\excel_plugin"
   ```

3. **If using a .pfx file** (from your organization or purchased):
   ```powershell
   $cert = Get-PfxCertificate -FilePath "C:\path\to\certificate.pfx"
   Set-AuthenticodeSignature -FilePath "chart_styles.xlam" -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
   ```
   You will be prompted to enter the certificate password.

4. **If using a certificate in the system store** (already installed on your machine):
   ```powershell
   $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
   Set-AuthenticodeSignature -FilePath "chart_styles.xlam" -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
   ```

   The `-TimestampServer` parameter adds a trusted timestamp so the signature remains valid even after the certificate expires.

5. Verify the signature was applied:
   ```powershell
   Get-AuthenticodeSignature -FilePath "chart_styles.xlam"
   ```
   You should see:
   ```
   SignerCertificate      Status   Path
   -----------------      ------   ----
   [certificate details]   Valid    chart_styles.xlam
   ```

#### Sign using Visual Studio (alternative)

If you have Visual Studio installed, you can sign from the command line:
```cmd
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com chart_styles.xlam
```

### Step 3: Verify the signature on a clean machine

To confirm the signature works, copy the signed `.xlam` to another Windows machine and check the signature:
1. Right-click the `.xlam` file → **Properties**
2. Click the **Digital Signatures** tab (if present, the signature is recognized)
3. Open PowerShell and run:
   ```powershell
   Get-AuthenticodeSignature -FilePath "chart_styles.xlam"
   ```

If your organization's certificate is from an internal CA, you may need to ensure the CA's root certificate is in the **Trusted Root Certification Authorities** store on user machines. Your IT team can distribute this via Group Policy (Active Directory).

---

## Part 2: Trust Center Configuration

### Why configure the Trust Center?

By default, Excel may block unsigned or untrusted add-ins. The Trust Center allows you to:
- Define which add-ins are permitted to run
- Require digital signatures
- Control notification levels

### Microsoft Trust Center documentation

- [Enable or disable macros in Microsoft Office](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-office-documents-12b043b5-4bac-46ee-37a0-6f04a7da7674)
- [Macro security and privacy settings](https://support.microsoft.com/en-us/office/macro-security-and-privacy-settings-in-excel-67f2e3ea-04e5-4ef9-8b52-81c5e2b91c78)

### For end users: Configure Trust Center manually

1. In Excel: **File → Options → Trust Center → Trust Center Settings**
2. Click **Trusted Locations** (left sidebar)
3. Click **Add New Location**
4. Browse to the folder where the `.xlam` file is stored and click **OK**
5. Repeat for all folders from which users might load the add-in

Alternatively, for **Trusted Publishers**:
1. **File → Options → Trust Center → Trust Center Settings → Trusted Publishers**
2. Ensure the certificate issuer (or root CA) is listed
3. If not, click **Add Publisher...** and browse to the signed `.xlam` file — Excel will extract and trust the issuer

### For IT administrators: Deploy via Group Policy (Active Directory)

If your organization uses Active Directory, you can push Trust Center settings to all machines:

1. **Create a registry file** that adds the add-in's folder to the Trusted Locations. Example for Windows Registry Editor Format 5.00:
   ```
   Windows Registry Editor Version 5.00

   [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations]
   "Location1"="\\\\company-server\\shared\\addins\\"
   "Path1"="\\\\company-server\\shared\\addins\\"
   "Date1"="2026-06-19"
   "Description1"="Company Chart Styles Add-in"
   ```

2. **Deploy via Group Policy:**
   - Open Group Policy Editor (`gpedit.msc` on domain-joined machines)
   - Navigate to **User Configuration → Preferences → Windows Settings → Registry**
   - Import the `.reg` file
   - Link the policy to the Organizational Unit (OU) containing your users

3. **Or use Intune** (for cloud-managed devices):
   - Upload the registry configuration as a custom profile
   - Assign to user groups
   - See [Manage devices with Intune](https://learn.microsoft.com/en-us/mem/intune/) for details

### For IT administrators: Trust Center via Office Cloud Policy (modern approach)

Microsoft now recommends **Office Cloud Policy Service** for managing Trust Center settings:

1. Go to [Office Cloud Policy Service](https://config.office.com)
2. Sign in with an admin account
3. Create a new policy configuration
4. Set **Macro Settings** to "Allow macros to run" or "Notify for unsigned macros"
5. Add your add-in's path to **Trusted Locations**
6. Deploy to users or groups
7. See [Overview of cloud policies for Microsoft 365 Apps](https://learn.microsoft.com/en-us/deployoffice/admincenter/overview-office-cloud-policy-service) for step-by-step instructions

---

## Part 3: Deployment Methods

### Method 1: Network share (simplest for small deployments)

1. Place the signed `.xlam` file on a **shared network drive** accessible to all users
   ```
   \\company-server\shared\addins\chart_styles.xlam
   ```

2. Each user adds the location as a Trusted Location in Excel:
   - **File → Options → Trust Center → Trusted Locations → Add New Location**
   - Enter the path to the shared folder
   - Click OK

3. Users load the add-in:
   - **File → Options → Add-ins → Manage Excel Add-ins → Browse**
   - Navigate to the shared folder and select `chart_styles.xlam`
   - Click OK

**Pros:** Simple, no IT infrastructure required, easy to update (just replace the file)
**Cons:** Requires manual setup per user, network dependency, doesn't scale to hundreds of users

---

### Method 2: Office Add-in Store (centralized, recommended for large organizations)

Organizations with Office 365/Microsoft 365 can publish add-ins to the **Office Add-in Store** for streamlined distribution.

1. **Package the add-in for the Store:**
   - Create a manifest file describing the add-in (see [Create your Office Add-in manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/create-addin-commands))
   - Prepare marketing materials (icon, description)

2. **Submit to the Office Add-in Store:**
   - Go to [Partner Center](https://partner.microsoft.com/en-us/dashboard)
   - Submit your add-in for certification
   - Microsoft validates the add-in (typically 5–7 business days)

3. **Deploy to your organization:**
   - In Microsoft 365 admin center, go to **Settings → Integrated apps**
   - Enable **Allow access to Office Add-ins** and **Allow user submissions**
   - Add your add-in from the Store
   - Assign to user groups

**Pros:** No manual user setup, automatic updates, Microsoft-curated trust
**Cons:** Requires certification, not suitable for internal-only add-ins

**See:** [Publish your Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish)

---

### Method 3: Centralized deployment (Microsoft 365 admin center, recommended for large organizations)

If your organization uses Microsoft 365, you can deploy add-ins directly from the admin center.

1. **Prepare the add-in manifest:**
   - Create a `manifest.xml` file describing your add-in
   - Host it on a public HTTPS endpoint (or provide the `.xlam` file directly)

2. **In Microsoft 365 Admin Center:**
   - Go to **Settings → Integrated apps → Deploy an app**
   - Choose **Upload a custom app** or **Add from Office Store**
   - Upload your manifest or select your app
   - Configure permissions
   - Assign to users or groups

3. **Users receive the add-in automatically** in their Office applications

**Pros:** Scales to thousands of users, no manual setup, IT-managed, automatic updates
**Cons:** Requires Microsoft 365 admin access, needs manifest file

**See:** [Centralized deployment in the Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage-deployment-of-add-ins)

---

### Method 4: Group Policy (for on-premises organizations without Office 365)

Organizations with Active Directory (not using Microsoft 365) can deploy via Group Policy.

1. **Create a registry entry** that auto-loads the add-in:
   ```
   [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Addins\chart_styles.xlam]
   "Path"="\\\\company-server\\shared\\addins\\chart_styles.xlam"
   "LoadBehavior"=dword:00000003
   ```
   (LoadBehavior = 3 means "load at startup")

2. **Deploy via Group Policy:**
   - Create a Group Policy Object (GPO) in Active Directory
   - Import the registry settings
   - Link to the Organizational Unit containing your users

3. **The add-in loads automatically** when users open Excel

**Pros:** No user intervention, automatic at startup
**Cons:** Requires Active Directory, on-premises only

**See:** [Deploy Office Add-ins using Group Policy](https://learn.microsoft.com/en-us/deployoffice/manage-add-ins-using-group-policy)

---

## Part 4: Access Control (Limiting to Specific User Groups)

### Scenario: You want only certain users or departments to access the add-in

#### Method 1: File system permissions (simplest)

If the add-in is on a network share, use **NTFS permissions**:

1. Right-click the folder containing the `.xlam` file → **Properties → Security**
2. Click **Edit** and configure permissions:
   - **Allow Read & Execute** for authorized users/groups
   - **Deny** for all other users (or remove them)
3. Click **Apply** and confirm

Only users with read access can load the add-in. This is simple but doesn't prevent users from finding the file elsewhere.

#### Method 2: Group Policy assignment (recommended for large organizations)

If deploying via Group Policy, assign the policy to specific Organizational Units (OUs):

1. Create a separate GPO for the add-in
2. In **Group Policy Editor**, edit the policy
3. Scope it to apply only to specific OUs or security groups
4. Only users in those groups receive the registry entries

**Example:** Create an OU for "Marketing Department" and link the add-in policy only to that OU.

See [Group Policy Filtering](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-defender-application-guard/reqs-wd-app-guard)

#### Method 3: Microsoft Entra ID / Conditional Access (cloud-first approach)

For organizations using Microsoft 365 with Entra ID (formerly Azure AD):

1. In **Microsoft 365 Admin Center**, assign the add-in to specific security groups:
   - Go to **Settings → Integrated apps → Manage apps**
   - Select your app
   - Click **Assign or unassign apps**
   - Choose the security groups to grant access

2. Only users in those groups see or can install the add-in

3. You can also use **Conditional Access policies** to require specific conditions (location, device compliance, etc.) before the add-in loads

**See:** [Assign Microsoft 365 apps to groups](https://learn.microsoft.com/en-us/microsoft-365/admin/manage-deployment-of-add-ins)

#### Method 4: Role-based access control within the add-in (code-level)

If you need fine-grained control, you can add code to the add-in that checks the current user against a whitelist:

**Example in `modRibbonHandlers.bas`:**
```vba
Private Function IsUserAuthorized() As Boolean
    Dim allowedUsers As Variant
    allowedUsers = Array("user1@company.com", "user2@company.com", "deptmarketing@company.com")
    
    Dim currentUser As String
    currentUser = Environ("USERNAME")  ' or use a more robust method with Office object model
    
    Dim i As Integer
    For i = LBound(allowedUsers) To UBound(allowedUsers)
        If currentUser = allowedUsers(i) Then
            IsUserAuthorized = True
            Exit Function
        End If
    Next i
    
    IsUserAuthorized = False
End Function

Public Sub CheckAuthorizationOnLoad()
    If Not IsUserAuthorized() Then
        MsgBox "You are not authorized to use this add-in. Please contact IT.", vbCritical
        ' Optionally disable the ribbon or exit
    End If
End Sub
```

Then call `CheckAuthorizationOnLoad()` when the add-in initializes (in `Workbook_Open`).

**Pros:** Granular control, can log usage
**Cons:** Requires code changes, less transparent to users

---

## Part 5: Versioning and Updates

### Managing updates safely

1. **Version the .xlam file:**
   Use a naming scheme like `chart_styles_v1.0.xlam`, `chart_styles_v1.1.xlam`, etc.

2. **Keep old versions available** for a grace period in case users need to revert.

3. **Announce updates** to users before rolling out:
   - Email a change log
   - Post to your intranet
   - Provide a deadline for updating

4. **Re-sign each version** with your code-signing certificate.

5. **Update centralized deployment:**
   - If using the Office Add-in Store or Microsoft 365 admin center, upload the new version
   - Users are notified automatically and can update at their convenience

6. **Update network shares:**
   - Replace the old `.xlam` with the new version
   - Keep a backup of the old version in a dated folder for rollback

---

## Part 6: Monitoring and Security

### Monitor add-in usage

1. **Office Add-in logging:**
   - Add logging to your add-in code (e.g., `Debug.Print` statements that write to a worksheet or log file)
   - When users encounter issues, they can share the log

2. **SharePoint document library telemetry:**
   - If users interact with SharePoint, audit logs track document access
   - See [Audit logging in Microsoft 365](https://learn.microsoft.com/en-us/microsoft-365/compliance/detailed-properties-in-the-audit-log)

3. **Optional: Application Insights:**
   - For more advanced telemetry, integrate [Application Insights](https://learn.microsoft.com/en-us/azure/azure-monitor/app/app-insights-overview) (requires Office JavaScript API, not typical for VBA add-ins)

### Revoke a signed certificate (if compromised)

If your code-signing certificate is compromised:

1. **Revoke the certificate** with your CA:
   - Contact your IT/Security team (for internal CA) or the CA vendor
   - Provide the certificate serial number
   - The CA will issue a Certificate Revocation List (CRL) update

2. **Sign future versions with a new certificate**

3. **Notify users** to disable or uninstall the old add-in

4. **Redeploy the new signed version** via your chosen method

---

## Checklist: Ready to Deploy

Before rolling out to users:

- [ ] The `.xlam` file is built and tested ([guide_building_xlam.md](guide_building_xlam.md))
- [ ] The `.xlam` file is digitally signed with a valid code-signing certificate
- [ ] The signature is verified on a clean machine
- [ ] Trust Center locations or policies are configured on test machines
- [ ] The add-in loads without errors in Excel
- [ ] A deployment method is chosen (network share, Group Policy, Microsoft 365 admin center, etc.)
- [ ] User access control is configured (file permissions, group assignment, etc.)
- [ ] Documentation is ready for end users (how to load, troubleshoot)
- [ ] IT team is prepared to support issues (configuration, revocation plan, etc.)
- [ ] Version numbering scheme is established

---

## Troubleshooting

| Problem | Likely cause | Solution |
|---|---|---|
| "This Office Add-in has been disabled" | Unsigned or untrusted add-in | Sign the add-in; configure Trust Center |
| Signature shows as "Invalid" or "Unknown" | Certificate not in trusted store | Ensure CA root cert is in **Trusted Root Certification Authorities** |
| Users can't find the add-in | Network path incorrect or permissions missing | Verify path in GPO/Policy; check NTFS permissions |
| Add-in doesn't load on startup | LoadBehavior registry value wrong | Check registry; ensure value is `dword:00000003` |
| Users report "access denied" | File system permissions too restrictive | Grant read access to authorized users/groups |

---

## Further Reading

- [Macro security and privacy settings in Excel](https://support.microsoft.com/en-us/office/macro-security-and-privacy-settings-in-excel-67f2e3ea-04e5-4ef9-8b52-81c5e2b91c78)
- [Code Signing Best Practices](https://learn.microsoft.com/en-us/previous-versions/dotnet/articles/ms998297(v=msdn.10))
- [Deploy Office Add-ins for Microsoft 365](https://learn.microsoft.com/en-us/deployoffice/overview-deploying-microsoft-365-apps)
- [Office Cloud Policy Service](https://learn.microsoft.com/en-us/deployoffice/admincenter/overview-office-cloud-policy-service)
- [Centralized deployment in the Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-365/admin/manage-deployment-of-add-ins)
- [Active Directory and Group Policy](https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-defender-application-guard/reqs-wd-app-guard)
