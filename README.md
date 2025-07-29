# HarjTech SharePoint Online Management Toolkit

## Overview

The **HarjTech SharePoint Online Management Toolkit** is a reusable PowerShell script designed to help SharePoint Online administrators streamline their daily tasks.  It leverages the [PnP.PowerShell](https://github.com/pnp/powershell) module to provide a rich set of functions for managing modern SharePoint sites, lists, items, columns, permissions and the recycle bin.  By consolidating common administrative actions into a single script, HarjTech reduces the time and risk associated with manual management through the SharePoint user interface.

### Why use this script?

* **Efficiency** – routine tasks such as creating sites, auditing lists, and modifying metadata are encapsulated into easy‑to‑call functions.  The script uses modern PnP commands like **`New‑PnPSite`**, which the official documentation states is used to create modern site collections and requires a mandatory `-Type` parameter.  Having these functions in one place reduces the need to search for individual commands and ensures consistency across administrators.
* **Accuracy** – each function wraps a PnP command with sensible parameter defaults and error handling.  For example, adding a column calls **`Add‑PnPField`**, which adds a field to a list or as a site column, while adding content types uses **`Add‑PnPContentTypeToList`**, allowing you to set the default content type.  Using these commands through a script helps avoid mis‑typed parameters and destructive mistakes.
* **Transparency** – comprehensive comments explain what each function does, why it matters, and how it ties back to SharePoint functionality.  Administrators can quickly understand the impact and purpose of each command.
* **Portability** – the script works on Windows, Linux, or macOS with PowerShell 7.x and the PnP.PowerShell module installed.  It can be dot‑sourced into any session or imported into existing automation.

## Requirements

* **PowerShell 7.x or later** – the script uses modern PowerShell features and should be run in a cross‑platform environment.
* **[PnP.PowerShell module](https://github.com/pnp/powershell)** – install via `Install‑Module PnP.PowerShell`.  Ensure it is up to date.
* **SharePoint Online permissions** – depending on the function, you may need:
  * **Global or SharePoint administrator** rights to create or delete site collections via `New‑PnPSite` and `Remove‑PnPTenantSite`.
  * **Site collection administrator** rights to manage lists, items, columns, permissions and recycle bin.  For example, retrieving recycle bin items requires site collection admin privileges and restoring them uses `Restore‑PnPRecycleBinItem`【248364638113375†L798-L832】.
* **Tenant admin URL** – to manage tenant‑wide operations (e.g. `GetAllSites` or `DeleteASite`), connect to the tenant admin URL (e.g. `https://tenant-admin.sharepoint.com`) using the `Connect‑HarjTechSPO` function.

## Usage

1. **Download and unzip the package.**  The zip file contains `harjtech_spo_management.ps1` and this `README.md`.
2. **Install PnP.PowerShell** (if not already):

   ```powershell
   Install‑Module PnP.PowerShell -Scope CurrentUser
   ```

3. **Dot‑source the script** in your PowerShell session to load the functions:

   ```powershell
   . .\harjtech_spo_management.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance"
   ```

   The `-SiteUrl` parameter sets the default SharePoint Online site for the session.  All subsequent functions assume that you have connected with `Connect‑HarjTechSPO`.

4. **Authenticate** using `Connect‑HarjTechSPO`:

   ```powershell
   Connect‑HarjTechSPO -SiteUrl "https://contoso.sharepoint.com/sites/Finance"
   ```

5. **Call functions as needed.**  For example, to list all lists on the connected site:

   ```powershell
   GetAllLists | Select Title, BaseType
   ```

## Included Functions

| Function | Description |
|---------|-------------|
| **Connect‑HarjTechSPO** | Authenticates to a SharePoint Online site or tenant admin site using interactive login and stores the connection for reuse. |
| **CreateNewSite** | Creates modern site collections (TeamSite, CommunicationSite or TeamSiteWithoutMicrosoft365Group).  As Microsoft notes, this cmdlet requires a `-Type` parameter specifying the site type. |
| **GetASite** | Retrieves properties of the current connection or a specified site. |
| **DeleteASite** | Deletes a site collection using `Remove‑PnPTenantSite` (requires tenant admin rights). |
| **GetAllSites** | Lists all site collections in the tenant. |
| **GetAllLists** | Retrieves all lists and libraries in the current site, optionally including hidden lists. |
| **GetAList** | Returns a specific list by name, ID or URL. |
| **GetAItem** | Retrieves a list item by ID, optionally selecting specific columns. |
| **GetAllItems** | Retrieves all items from a list with optional field selection and CAML query filtering. |
| **AddAColumnToaList** | Adds a column to a list using `Add‑PnPField`, which adds a field (column) to a list or site column. |
| **RemoveAColumnToAList** | Removes a column from a list with `Remove‑PnPField`. |
| **UpdateAColumnForSingleItem** | Updates specified column values on a single list item. |
| **UpdateAColumnForMultipleItems** | Bulk updates items by ID array or CAML query. |
| **AddContentTypeToList** | Adds an existing content type to a list and optionally sets it as default. |
| **GetAllContentTypesBeingUsedOnAList** | Lists all content types associated with a list, using `Get‑PnPContentType`. |
| **GetAllPermissionLevelOnASite** | Retrieves all role definitions (permission levels) for the current site. |
| **GetAllSPOSitePermissions** | Produces a report of SharePoint groups and their members with assigned roles. |
| **CheckIfAUserHaveAccessToASite** | Heuristically determines whether a user has access to the current site. |
| **CheckIfAUserHaveAccessToAList** | Checks a user’s list permissions via `Get‑PnPListPermissions'. |
| **CheckIfAUserHaveAccessToAListItem** | Determines if a user has permissions on a specific list item using `Get‑PnPListItemPermission`. |
| **CheckIfAuserIsinaSharePointGroup** | Verifies group membership for a user. |
| **GetAllSPOSiteGroups** | Lists all SharePoint groups defined on the current site. |
| **AddUserToaSPOSIteGroup** | Adds a user to a SharePoint group. |
| **RemoveUserFromASiteGroup** | Removes a user from a SharePoint group. |
| **RestoreAItemFromRecycleBin** | Restores items from the recycle bin using `Restore‑PnPRecycleBinItem`. |
| **CheckIfFileOrListItemIsWithinTheRecycleBin** | Searches the recycle bin for deleted files or list items; requires site collection admin privileges. |

### Parameter Conventions

Most functions accept parameters such as `List`, `Identity`, `ItemId`, `Values`, `LoginName` and `GroupName`.  Names are self‑explanatory; see the script comments for detailed information, examples and optional parameters.  Where appropriate, functions allow passing either names or GUIDs.

## Tenant and permission considerations

* **Site creation and deletion** require connecting to the tenant admin URL and running the session as a SharePoint or global administrator.
* **List and item operations** require at least site collection administrator privileges on the target site.
* **Recycle bin functions** require site collection admin rights.  `Get‑PnPRecycleBinItem` returns all items in the recycle bin and is limited to site collection admins.

## Marketing note

This toolkit was engineered by **HarjTech Solutions** as part of our commitment to enhancing your cloud productivity.  We specialize in SharePoint architecture, governance and automation.  If you find this script useful, check out our other solutions at:

* **Website:** [www.harjtech.com](https://www.harjtech.com) – explore our consulting services, training programs and custom development offerings.
* **GitHub:** [HarjTechInc on GitHub](https://github.com/HarjTechInc) – discover additional open‑source tools and contributions.

We welcome your feedback and collaboration.  Together we can modernize your SharePoint environment and accelerate your business growth.
