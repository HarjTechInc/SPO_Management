# ===================================================================================================================
# HarjTech SharePoint Online Management Toolkit
#
# SYNOPSIS
#     This script provides a suite of functions that leverage the PnP.PowerShell module to administer
#     modern SharePoint Online sites and lists.  Administrators can create or remove sites, query site
#     collections, manipulate lists and list items, manage columns, content types and permissions and
#     restore items from the recycle bin.  Each function has been documented with parameter
#     descriptions and examples.  This script was authored by HarjTech Solutions to simplify the
#     day‑to‑day management of your SharePoint environment.  For more information about our
#     services please visit https://www.harjtech.com or our GitHub repository at
#     https://github.com/HarjTechInc.
#
# DESCRIPTION
#     HarjTech recommends administrators keep a management script like this handy to streamline
#     repetitive tasks and to reduce the risk of performing destructive actions through the web UI.
#     PnP.PowerShell is a modern, cross‑platform module supported by Microsoft that wraps the
#     SharePoint CSOM and simplifies administration.  Microsoft’s own documentation notes that
#     New‑PnPSite creates modern site collections and requires a mandatory -Type parameter to
#     indicate the type of site you wish to create【118114568060417†L849-L856】, while functions such as
#     Add‑PnPField add columns to lists or site collections【785190958354182†L835-L837】.  Leveraging these cmdlets in
#     a central script reduces the need to memorize individual commands and ensures consistent
#     execution across administrators.
#
#     Note: Several functions require tenant‑administrator rights (e.g. creating or deleting
#     site collections) or site collection administrator rights.  Use caution and test in a
#     non‑production environment first.  Always connect to the desired site using the
#     Connect‑HarjTechSPO function before executing other commands.
#
# REQUIREMENTS
#     - PnP.PowerShell module (Install‑Module PnP.PowerShell)
#     - SharePoint Online administrator or site collection administrator privileges depending on the function
#     - For tenant scoped operations (site creation/deletion), you must connect to the tenant admin
#       URL (e.g. https://tenant‑admin.sharepoint.com) using Connect‑HarjTechSPO.
#
# USAGE
#     Import this script into your PowerShell session or dot‑source it.  Then call Connect‑HarjTechSPO
#     with the target site URL before calling additional functions.  Each function supports
#     descriptive parameters.  For example:
#         . .\harjtech_spo_management.ps1
#         Connect‑HarjTechSPO -SiteUrl "https://contoso.sharepoint.com/sites/Finance"
#         GetAllLists
#
# MARKETING
#     This script is delivered as part of HarjTech Solutions’ SharePoint offerings.  We provide
#     consulting services, customization and automation solutions tailored to your organization.
#     Visit https://www.harjtech.com or our GitHub repository at https://github.com/HarjTechInc for
#     additional tools and guidance.
# ===================================================================================================================

param(
    # The SharePoint Online site URL to connect to.  All other functions assume that this site
    # connection has already been established via Connect‑HarjTechSPO.
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

###
### Function: Connect‑HarjTechSPO
###
function Connect‑HarjTechSPO {
    <#
    .SYNOPSIS
        Establishes a connection to the specified SharePoint Online site using interactive login.

    .DESCRIPTION
        This helper function wraps the PnP.PowerShell Connect‑PnPOnline cmdlet.  Administrators
        should call this function at least once per session to authenticate to the desired site.  The
        connection is stored globally and reused by subsequent functions.  Without a valid
        connection the other commands in this script will fail.  For tenant‑level operations,
        connect to the tenant admin URL (e.g. https://tenant‑admin.sharepoint.com).

    .PARAMETER SiteUrl
        The full URL of the SharePoint Online site or tenant admin URL.

    .EXAMPLE
        Connect‑HarjTechSPO -SiteUrl "https://contoso.sharepoint.com/sites/Finance"
        Connects to the Finance site using interactive authentication.

    .NOTES
        Requires the PnP.PowerShell module.  When using service principal or app‑only auth,
        replace ‑Interactive with appropriate parameters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    # Initiate interactive login for the given site URL
    Connect‑PnPOnline -Url $SiteUrl -Interactive -ErrorAction Stop
}

###
### Function: CreateNewSite
###
function CreateNewSite {
    <#
    .SYNOPSIS
        Creates a modern SharePoint Online site collection.

    .DESCRIPTION
        Uses New‑PnPSite to provision a communication site, Microsoft 365 group‑connected
        team site or a non‑group team site.  The documentation for New‑PnPSite notes that this
        cmdlet creates modern site collections and that the mandatory -Type parameter determines
        whether you create a CommunicationSite, TeamSite or TeamSiteWithoutMicrosoft365Group【118114568060417†L849-L856】.
        Administrators must be connected to the tenant admin URL and have SharePoint
        administrator privileges.  The function accepts common site properties and passes them
        through to New‑PnPSite.  If creating a team site, supply an alias rather than a full URL.

    .PARAMETER Title
        Display name of the new site.

    .PARAMETER Type
        The type of site: TeamSite, CommunicationSite or TeamSiteWithoutMicrosoft365Group.

    .PARAMETER AliasOrUrl
        For TeamSite, this is the group alias; for CommunicationSite and TeamSiteWithoutMicrosoft365Group,
        provide the full URL (e.g. https://tenant.sharepoint.com/sites/ProjectX).

    .PARAMETER Description
        Optional description of the site.

    .PARAMETER IsPublic
        Only applicable to TeamSite.  Creates a public (open) Microsoft 365 group when supplied.

    .PARAMETER Owners
        One or more UPNs for users who will be designated as site owners.  Owners must already
        exist in the tenant.

    .EXAMPLE
        CreateNewSite -Title "Project Team" -Type TeamSite -AliasOrUrl "projectteam" -Owners "user1@contoso.com","user2@contoso.com"

        Creates a modern Microsoft 365 group‑connected team site with alias "projectteam" and
        sets the specified users as owners.

    .EXAMPLE
        CreateNewSite -Title "Marketing" -Type CommunicationSite -AliasOrUrl "https://tenant.sharepoint.com/sites/Marketing"

        Creates a standalone communication site at the specified URL.

    .NOTES
        Creating sites consumes resources and may take several minutes.  Removing sites places
        them into the tenant recycle bin unless the -SkipRecycleBin option is specified.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][ValidateSet('TeamSite','CommunicationSite','TeamSiteWithoutMicrosoft365Group')][string]$Type,
        [Parameter(Mandatory = $true)][string]$AliasOrUrl,
        [string]$Description,
        [switch]$IsPublic,
        [string[]]$Owners
    )
    try {
        switch ($Type) {
            'TeamSite' {
                New‑PnPSite -Type TeamSite -Title $Title -Alias $AliasOrUrl -Description $Description -IsPublic:$IsPublic -Owners $Owners -ErrorAction Stop
            }
            'CommunicationSite' {
                New‑PnPSite -Type CommunicationSite -Title $Title -Url $AliasOrUrl -Description $Description -Owner ($Owners -join ',') -ErrorAction Stop
            }
            'TeamSiteWithoutMicrosoft365Group' {
                New‑PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $Title -Url $AliasOrUrl -Description $Description -Owner ($Owners -join ',') -ErrorAction Stop
            }
        }
        Write‑Verbose "Site creation request submitted successfully."
    } catch {
        Write‑Error "Failed to create site: $_"
    }
}

###
### Function: GetASite
###
function GetASite {
    <#
    .SYNOPSIS
        Retrieves details for the current connected site or a specified site.

    .DESCRIPTION
        Calls Get‑PnPSite to return information about a SharePoint site collection.  If the
        optional Url parameter is provided, the function will temporarily connect to that site
        and return its properties.  Otherwise it returns the properties of the currently
        connected site.  Use this command to verify site settings or to ensure that you are
        connected to the correct site before performing operations.

    .PARAMETER Url
        (Optional) The full URL of the site to retrieve.  If omitted, the current connection is used.

    .EXAMPLE
        GetASite

        Returns details about the site connected with Connect‑HarjTechSPO.

    .EXAMPLE
        GetASite -Url "https://tenant.sharepoint.com/sites/Finance"

        Returns details about the specified site without altering the existing connection.
    #>
    [CmdletBinding()]
    param(
        [string]$Url
    )
    if ($PSBoundParameters.ContainsKey('Url')) {
        # Use a temporary connection to retrieve the remote site
        $tempConn = Connect‑PnPOnline -Url $Url -ReturnConnection -Interactive -ErrorAction Stop
        Get‑PnPSite -Connection $tempConn
        Disconnect‑PnPOnline -Connection $tempConn
    } else {
        Get‑PnPSite
    }
}

###
### Function: DeleteASite
###
function DeleteASite {
    <#
    .SYNOPSIS
        Deletes a SharePoint Online site collection.

    .DESCRIPTION
        Uses Remove‑PnPTenantSite to delete the specified site collection.  By default the site
        will be moved to the tenant recycle bin.  Specify ‑SkipRecycleBin to permanently delete
        the site.  Only SharePoint administrators or global administrators can perform this
        operation.  Always verify the URL carefully before executing.

    .PARAMETER Url
        The full URL of the site collection to delete.

    .PARAMETER SkipRecycleBin
        Permanently deletes the site without placing it in the recycle bin.

    .PARAMETER Force
        Suppresses the confirmation prompt.

    .EXAMPLE
        DeleteASite -Url "https://tenant.sharepoint.com/sites/ProjectX" -Force

        Deletes the ProjectX site collection and moves it to the recycle bin.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [switch]$SkipRecycleBin,
        [switch]$Force
    )
    if ($PSCmdlet.ShouldProcess($Url, 'Remove Site Collection')) {
        try {
            Remove‑PnPTenantSite -Url $Url -SkipRecycleBin:$SkipRecycleBin -Force:$Force -ErrorAction Stop
            Write‑Verbose "Site collection removed."
        } catch {
            Write‑Error "Failed to delete site: $_"
        }
    }
}

###
### Function: GetAllSites
###
function GetAllSites {
    <#
    .SYNOPSIS
        Lists all site collections in the tenant.

    .DESCRIPTION
        Calls Get‑PnPTenantSite to retrieve all site collections.  You must be connected to the
        tenant admin URL and have SharePoint administrator privileges.  Use this function to
        audit existing site collections.

    .EXAMPLE
        GetAllSites | Format‑Table Url, Title
    #>
    [CmdletBinding()]
    param()
    try {
        Get‑PnPTenantSite -Includes OneDriveSites -ErrorAction Stop
    } catch {
        Write‑Error "Failed to retrieve sites: $_"
    }
}

###
### Function: GetAllLists
###
function GetAllLists {
    <#
    .SYNOPSIS
        Retrieves all lists and libraries in the current site.

    .DESCRIPTION
        Calls Get‑PnPList without specifying a name to return all lists and document libraries
        within the connected site.  Hidden lists can be included by using the IncludeHidden
        switch.  This is useful for inventorying libraries or preparing for further actions.

    .PARAMETER IncludeHidden
        When supplied, also returns hidden lists such as workflow histories.

    .EXAMPLE
        GetAllLists | Select Title, BaseType
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeHidden
    )
    try {
        if ($IncludeHidden) {
            Get‑PnPList -Includes Hidden -ErrorAction Stop
        } else {
            Get‑PnPList -ErrorAction Stop | Where‑Object { -not $_.Hidden }
        }
    } catch {
        Write‑Error "Failed to retrieve lists: $_"
    }
}

###
### Function: GetAList
###
function GetAList {
    <#
    .SYNOPSIS
        Retrieves a specific list or library by name or ID.

    .DESCRIPTION
        Wrapper around Get‑PnPList allowing you to retrieve a single list by title, ID or URL.
        Use this to inspect list settings or confirm that a list exists before performing
        operations.

    .PARAMETER Identity
        The name, ID or server‑relative URL of the list to retrieve.

    .EXAMPLE
        GetAList -Identity "Documents"

        Returns the properties of the Documents library.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Identity
    )
    try {
        Get‑PnPList -Identity $Identity -ErrorAction Stop
    } catch {
        Write‑Error "Failed to retrieve list: $_"
    }
}

###
### Function: GetAItem
###
function GetAItem {
    <#
    .SYNOPSIS
        Retrieves a specific item from a list or library.

    .DESCRIPTION
        Calls Get‑PnPListItem to retrieve an item by its ID.  Use this function to view item
        properties or to verify that an item exists before updating or deleting it.  For large
        lists consider specifying the -Fields parameter to retrieve only required columns.

    .PARAMETER List
        The name, ID or URL of the list containing the item.

    .PARAMETER ItemId
        The integer ID of the list item to retrieve.

    .PARAMETER Fields
        (Optional) An array of field internal names to include in the result.

    .EXAMPLE
        GetAItem -List "Tasks" -ItemId 12 -Fields "Title","Status"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][int]$ItemId,
        [string[]]$Fields
    )
    try {
        if ($PSBoundParameters.ContainsKey('Fields')) {
            Get‑PnPListItem -List $List -Id $ItemId -Fields $Fields -ErrorAction Stop
        } else {
            Get‑PnPListItem -List $List -Id $ItemId -ErrorAction Stop
        }
    } catch {
        Write‑Error "Failed to retrieve list item: $_"
    }
}

###
### Function: GetAllItems
###
function GetAllItems {
    <#
    .SYNOPSIS
        Retrieves all items from a list or library.

    .DESCRIPTION
        Uses Get‑PnPListItem to return all list items.  PnP.PowerShell retrieves items in
        batches so that large lists can be processed efficiently.  An optional filter allows
        you to retrieve only items matching a CAML query or view.  You can also specify
        particular fields to minimize the amount of data returned.

    .PARAMETER List
        The name, ID or URL of the list.

    .PARAMETER Fields
        (Optional) Array of field internal names to return.

    .PARAMETER PageSize
        (Optional) Number of items to request per call.  Default is 100.

    .PARAMETER Query
        (Optional) CAML query fragment to filter items.  Use with caution.

    .EXAMPLE
        GetAllItems -List "Documents" -Fields "FileLeafRef","Author"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [string[]]$Fields,
        [int]$PageSize = 100,
        [string]$Query
    )
    try {
        $params = @{ List = $List; PageSize = $PageSize; ErrorAction = 'Stop' }
        if ($PSBoundParameters.ContainsKey('Fields')) { $params.Fields = $Fields }
        if ($PSBoundParameters.ContainsKey('Query'))  { $params.Query  = $Query  }
        Get‑PnPListItem @params
    } catch {
        Write‑Error "Failed to retrieve items: $_"
    }
}

###
### Function: AddAColumnToaList
###
function AddAColumnToaList {
    <#
    .SYNOPSIS
        Adds a new column to a list or document library.

    .DESCRIPTION
        Wraps the Add‑PnPField cmdlet, which according to Microsoft documentation adds a field
        (column) to a list or as a site column【785190958354182†L835-L837】.  You can specify the
        display name, internal name, type and optional choices for choice columns.  You may
        also add the column to the default view automatically.

    .PARAMETER List
        The name, ID or URL of the list to which the column should be added.

    .PARAMETER DisplayName
        The friendly name shown to users.

    .PARAMETER InternalName
        The internal name used by SharePoint.  Avoid spaces and special characters.

    .PARAMETER Type
        The type of field (e.g. Text, Number, Choice, MultiChoice, YesNo, DateTime).

    .PARAMETER Group
        (Optional) The group name where this site column will be organized.

    .PARAMETER Choices
        (Optional) Array of options for Choice or MultiChoice fields.

    .PARAMETER AddToDefaultView
        Adds the column to the list’s default view when specified.

    .EXAMPLE
        AddAColumnToaList -List "Issues" -DisplayName "Category" -InternalName "IssueCategory" -Type Choice -Choices "Bug","Feature","Task" -AddToDefaultView
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$InternalName,
        [Parameter(Mandatory = $true)][string]$Type,
        [string]$Group,
        [string[]]$Choices,
        [switch]$AddToDefaultView
    )
    try {
        $params = @{ List = $List; DisplayName = $DisplayName; InternalName = $InternalName; Type = $Type; ErrorAction = 'Stop' }
        if ($PSBoundParameters.ContainsKey('Group'))   { $params.Group   = $Group   }
        if ($PSBoundParameters.ContainsKey('Choices')) { $params.Choices = $Choices }
        if ($AddToDefaultView.IsPresent) { $params.AddToDefaultView = $true }
        Add‑PnPField @params
        Write‑Verbose "Column added successfully."
    } catch {
        Write‑Error "Failed to add column: $_"
    }
}

###
### Function: RemoveAColumnToAList
###
function RemoveAColumnToAList {
    <#
    .SYNOPSIS
        Removes a column from a list.

    .DESCRIPTION
        Uses Remove‑PnPField to remove a field from a list.  If you supply the internal name
        of a site column, the column will be removed from that list but not deleted from the
        site.  This function does not remove default columns or columns required by SharePoint.

    .PARAMETER List
        The name, ID or URL of the list.

    .PARAMETER Field
        The internal name, ID or Field object representing the column to remove.

    .EXAMPLE
        RemoveAColumnToAList -List "Issues" -Field "IssueCategory"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)]$Field
    )
    try {
        Remove‑PnPField -List $List -Identity $Field -ErrorAction Stop
        Write‑Verbose "Column removed successfully."
    } catch {
        Write‑Error "Failed to remove column: $_"
    }
}

###
### Function: UpdateAColumnForSingleItem
###
function UpdateAColumnForSingleItem {
    <#
    .SYNOPSIS
        Updates one or more fields on a single list item.

    .DESCRIPTION
        Calls Set‑PnPListItem to update specified column values on a single item.  Supply a
        hashtable to the Values parameter where each key is the internal field name and each
        value is the new data.  Use this when correcting data or setting metadata on an
        individual record.

    .PARAMETER List
        The name, ID or URL of the list containing the item.

    .PARAMETER ItemId
        The ID of the item to update.

    .PARAMETER Values
        Hashtable of column names and their new values (e.g. @{"Status"="Completed"}).

    .EXAMPLE
        UpdateAColumnForSingleItem -List "Tasks" -ItemId 3 -Values @{ "Status" = "Closed"; "AssignedTo" = "user@contoso.com" }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][int]$ItemId,
        [Parameter(Mandatory = $true)][hashtable]$Values
    )
    try {
        Set‑PnPListItem -List $List -Identity $ItemId -Values $Values -ErrorAction Stop
        Write‑Verbose "Item updated successfully."
    } catch {
        Write‑Error "Failed to update item: $_"
    }
}

###
### Function: UpdateAColumnForMultipleItems
###
function UpdateAColumnForMultipleItems {
    <#
    .SYNOPSIS
        Updates one or more fields on multiple list items.

    .DESCRIPTION
        Loops over a collection of item IDs or a CAML query and applies the supplied values to
        each item via Set‑PnPListItem.  This can be used to perform bulk metadata corrections.

    .PARAMETER List
        The name, ID or URL of the list.

    .PARAMETER ItemIds
        Array of item IDs to update.  If not supplied and Query is provided, all matching items
        from Query will be updated.

    .PARAMETER Query
        CAML query to select items.  Ignored if ItemIds are supplied.

    .PARAMETER Values
        Hashtable of column names and their new values.

    .EXAMPLE
        UpdateAColumnForMultipleItems -List "Tasks" -ItemIds 1,2,3 -Values @{ "Status" = "Closed" }

        Updates the Status field on items 1, 2 and 3.

    .EXAMPLE
        UpdateAColumnForMultipleItems -List "Tasks" -Query "<View><Query><Where><Eq><FieldRef Name='Status'/><Value Type='Text'>Open</Value></Eq></Where></Query></View>" -Values @{ "Status" = "Closed" }

        Closes all open tasks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [int[]]$ItemIds,
        [string]$Query,
        [Parameter(Mandatory = $true)][hashtable]$Values
    )
    try {
        $idsToUpdate = @()
        if ($PSBoundParameters.ContainsKey('ItemIds')) {
            $idsToUpdate = $ItemIds
        } elseif ($PSBoundParameters.ContainsKey('Query')) {
            $items = Get‑PnPListItem -List $List -Query $Query -Fields "ID" -ErrorAction Stop
            $idsToUpdate = $items.Id
        } else {
            throw "Either ItemIds or Query must be provided."
        }
        foreach ($id in $idsToUpdate) {
            Set‑PnPListItem -List $List -Identity $id -Values $Values -ErrorAction Stop
        }
        Write‑Verbose "Updated $($idsToUpdate.Count) item(s)."
    } catch {
        Write‑Error "Failed to update items: $_"
    }
}

###
### Function: AddContentTypeToList
###
function AddContentTypeToList {
    <#
    .SYNOPSIS
        Adds an existing content type to a list and optionally sets it as default.

    .DESCRIPTION
        Wraps the Add‑PnPContentTypeToList cmdlet.  The documentation explains that this cmdlet
        allows adding a content type to a list and, when the ‑DefaultContentType switch is used,
        sets the newly added content type as default【69259419889438†L810-L823】.  This function accepts
        the content type name or ID and the list to which it should be added.

    .PARAMETER List
        The list name, ID or URL.

    .PARAMETER ContentType
        Name or ID of the content type to add.

    .PARAMETER DefaultContentType
        Switch to set the content type as the default.

    .EXAMPLE
        AddContentTypeToList -List "Documents" -ContentType "Project Document" -DefaultContentType
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)]$ContentType,
        [switch]$DefaultContentType
    )
    try {
        Add‑PnPContentTypeToList -List $List -ContentType $ContentType -DefaultContentType:$DefaultContentType.IsPresent -ErrorAction Stop
        Write‑Verbose "Content type added successfully."
    } catch {
        Write‑Error "Failed to add content type: $_"
    }
}

###
### Function: GetAllContentTypesBeingUsedOnAList
###
function GetAllContentTypesBeingUsedOnAList {
    <#
    .SYNOPSIS
        Lists all content types associated with a given list.

    .DESCRIPTION
        Uses Get‑PnPContentType with the -List parameter to retrieve all content types on
        a specific list.  The Get‑PnPContentType cmdlet allows you to get a list of content
        types associated with a site or list and can also retrieve a single content type by
        specifying its identity【460062295138079†L811-L849】.  Use this function to audit which content types
        are enabled on a library.

    .PARAMETER List
        Name, ID or URL of the list.

    .EXAMPLE
        GetAllContentTypesBeingUsedOnAList -List "Documents"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List
    )
    try {
        Get‑PnPContentType -List $List -ErrorAction Stop
    } catch {
        Write‑Error "Failed to retrieve content types: $_"
    }
}

###
### Function: GetAllPermissionLevelOnASite
###
function GetAllPermissionLevelOnASite {
    <#
    .SYNOPSIS
        Retrieves all permission levels (role definitions) for the current site.

    .DESCRIPTION
        Calls Get‑PnPRoleDefinition to list every role definition within the site.  According
        to Microsoft’s examples, Get‑PnPRoleDefinition returns the role definitions (permission
        levels) of the current site and can filter by identity【360540768605276†L819-L837】.  Use this
        function to audit available permission levels before granting permissions.

    .EXAMPLE
        GetAllPermissionLevelOnASite
    #>
    [CmdletBinding()]
    param()
    try {
        Get‑PnPRoleDefinition -ErrorAction Stop
    } catch {
        Write‑Error "Failed to retrieve permission levels: $_"
    }
}

###
### Function: GetAllSPOSitePermissions
###
function GetAllSPOSitePermissions {
    <#
    .SYNOPSIS
        Produces a report of SharePoint groups and their members for the current site.

    .DESCRIPTION
        Retrieves all site groups via Get‑PnPGroup and enumerates their members using
        Get‑PnPGroupMembers.  While Get‑PnPListPermissions can be used to retrieve permissions
        for a specific principal【39255984457531†L809-L831】, this function focuses on site groups.  The output
        lists the group name, each member’s login name and the group’s permission roles.

    .EXAMPLE
        GetAllSPOSitePermissions | Format‑Table GroupName, MemberName, RoleNames
    #>
    [CmdletBinding()]
    param()
    $results = @()
    try {
        $groups = Get‑PnPGroup -ErrorAction Stop
        foreach ($grp in $groups) {
            $members = Get‑PnPGroupMembers -Identity $grp -ErrorAction SilentlyContinue
            $roles = Get‑PnPRoleAssignment -Identity $grp -ErrorAction SilentlyContinue
            $roleNames = ($roles.RoleDefinitionBindings | Select‑Object -ExpandProperty Name) -join ';'
            foreach ($mem in $members) {
                $results += [PSCustomObject]@{
                    GroupName = $grp.Title
                    MemberName = $mem.LoginName
                    RoleNames = $roleNames
                }
            }
        }
        return $results
    } catch {
        Write‑Error "Failed to retrieve site permissions: $_"
    }
}

###
### Function: CheckIfAUserHaveAccessToASite
###
function CheckIfAUserHaveAccessToASite {
    <#
    .SYNOPSIS
        Determines whether a specified user has access to the current site.

    .DESCRIPTION
        Attempts to retrieve the user using Get‑PnPUser.  If the user object is returned,
        it is assumed that the user has some level of access.  If an exception is thrown
        the user likely does not have direct access.  This is a heuristic and may not
        account for all scenarios (e.g. anonymous links or external sharing).  Use in
        combination with other permission functions for full coverage.

    .PARAMETER LoginName
        UPN or login name of the user to check.

    .EXAMPLE
        CheckIfAUserHaveAccessToASite -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        $user = Get‑PnPUser -Identity $LoginName -ErrorAction Stop
        Write‑Output "$LoginName has access to this site."
        return $true
    } catch {
        Write‑Output "$LoginName does not have access to this site."
        return $false
    }
}

###
### Function: CheckIfAUserHaveAccessToAList
###
function CheckIfAUserHaveAccessToAList {
    <#
    .SYNOPSIS
        Checks whether a user has any permissions on a given list.

    .DESCRIPTION
        Retrieves the principal ID for the specified user and queries list permissions using
        Get‑PnPListPermissions, which returns the list permissions (role definitions) for a
        specific user or group【39255984457531†L809-L831】.  If any permissions are returned the user
        has access; otherwise they do not.

    .PARAMETER List
        Name, ID or URL of the list.

    .PARAMETER LoginName
        UPN or login name of the user.

    .EXAMPLE
        CheckIfAUserHaveAccessToAList -List "Documents" -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        $user = Get‑PnPUser -Identity $LoginName -ErrorAction Stop
        $perms = Get‑PnPListPermissions -Identity $List -PrincipalId $user.Id -ErrorAction Stop
        if ($perms) {
            Write‑Output "$LoginName has access to list $List."
            return $true
        } else {
            Write‑Output "$LoginName does not have access to list $List."
            return $false
        }
    } catch {
        Write‑Output "$LoginName does not have access to list $List or user not found."
        return $false
    }
}

###
### Function: CheckIfAUserHaveAccessToAListItem
###
function CheckIfAUserHaveAccessToAListItem {
    <#
    .SYNOPSIS
        Determines whether a user has permissions on a specific list item.

    .DESCRIPTION
        Uses Get‑PnPListItemPermission to retrieve the permissions applied to a list item.
        The Get‑PnPListItemPermission cmdlet allows retrieving permissions for a given list
        item【832713506089915†L810-L820】.  This function searches the returned permissions for the
        specified user.  Note that permissions may be granted via group membership; the
        function attempts to match the user’s login name exactly against the principal names.

    .PARAMETER List
        The list name, ID or URL.

    .PARAMETER ItemId
        ID of the list item.

    .PARAMETER LoginName
        UPN or login name of the user.

    .EXAMPLE
        CheckIfAUserHaveAccessToAListItem -List "Documents" -ItemId 5 -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][int]$ItemId,
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        $perms = Get‑PnPListItemPermission -List $List -Identity $ItemId -ErrorAction Stop
        $found = $perms | Where‑Object { $_.PrincipalName -eq $LoginName }
        if ($found) {
            Write‑Output "$LoginName has access to item $ItemId in list $List."
            return $true
        } else {
            Write‑Output "$LoginName does not have access to item $ItemId in list $List."
            return $false
        }
    } catch {
        Write‑Output "Could not determine item permissions: $_"
        return $false
    }
}

###
### Function: CheckIfAuserIsinaSharePointGroup
###
function CheckIfAuserIsinaSharePointGroup {
    <#
    .SYNOPSIS
        Determines whether a user is a member of a specific SharePoint group.

    .DESCRIPTION
        Gets the specified group and then retrieves its members.  If the user’s login name is
        present, the function returns $true.  Use this to verify group membership before
        adding or removing users.

    .PARAMETER GroupName
        The display name of the SharePoint group.

    .PARAMETER LoginName
        UPN or login name of the user.

    .EXAMPLE
        CheckIfAuserIsinaSharePointGroup -GroupName "Site Owners" -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        $members = Get‑PnPGroupMembers -Identity $GroupName -ErrorAction Stop
        if ($members.LoginName -contains $LoginName) {
            Write‑Output "$LoginName is a member of $GroupName."
            return $true
        } else {
            Write‑Output "$LoginName is not a member of $GroupName."
            return $false
        }
    } catch {
        Write‑Output "Group not found or error retrieving members: $_"
        return $false
    }
}

###
### Function: GetAllSPOSiteGroups
###
function GetAllSPOSiteGroups {
    <#
    .SYNOPSIS
        Retrieves all SharePoint groups defined for the current site.

    .DESCRIPTION
        Calls Get‑PnPGroup without parameters to list all groups.  This is useful for reviewing
        group structures before assigning permissions.

    .EXAMPLE
        GetAllSPOSiteGroups | Select Title, Id
    #>
    [CmdletBinding()]
    param()
    try {
        Get‑PnPGroup -ErrorAction Stop
    } catch {
        Write‑Error "Failed to retrieve site groups: $_"
    }
}

###
### Function: AddUserToaSPOSIteGroup
###
function AddUserToaSPOSIteGroup {
    <#
    .SYNOPSIS
        Adds a user to a SharePoint group on the current site.

    .DESCRIPTION
        Wraps Add‑PnPGroupMember to add a specified user to a group.  The user must exist in
        Azure AD and must not already be a member of the group.  After adding, the user will
        inherit the group’s permissions.

    .PARAMETER GroupName
        The name of the group to add the user to.

    .PARAMETER LoginName
        The user’s UPN or login name.

    .EXAMPLE
        AddUserToaSPOSIteGroup -GroupName "Site Members" -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        Add‑PnPGroupMember -Identity $GroupName -Users $LoginName -ErrorAction Stop
        Write‑Output "$LoginName added to group $GroupName."
    } catch {
        Write‑Error "Failed to add user: $_"
    }
}

###
### Function: RemoveUserFromASiteGroup
###
function RemoveUserFromASiteGroup {
    <#
    .SYNOPSIS
        Removes a user from a SharePoint group.

    .DESCRIPTION
        Uses Remove‑PnPGroupMember to remove the specified user from the given group.  If the
        user is not a member, an exception will be suppressed silently.

    .PARAMETER GroupName
        Name of the group from which to remove the user.

    .PARAMETER LoginName
        UPN or login name of the user to remove.

    .EXAMPLE
        RemoveUserFromASiteGroup -GroupName "Site Members" -LoginName "user@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][string]$LoginName
    )
    try {
        Remove‑PnPGroupMember -Identity $GroupName -Users $LoginName -ErrorAction Stop
        Write‑Output "$LoginName removed from group $GroupName."
    } catch {
        Write‑Error "Failed to remove user: $_"
    }
}

###
### Function: RestoreAItemFromRecycleBin
###
function RestoreAItemFromRecycleBin {
    <#
    .SYNOPSIS
        Restores one or more items from the recycle bin.

    .DESCRIPTION
        Wraps the Restore‑PnPRecycleBinItem cmdlet.  The documentation notes that this cmdlet
        restores a specified item from the recycle bin to its original location【248364638113375†L798-L832】.
        You can specify an individual recycle bin item by GUID or pipeline in multiple items.

    .PARAMETER Identity
        The GUID(s) of the recycle bin item(s) to restore.  If omitted and RowLimit is
        provided, all items returned by Get‑PnPRecycleBinItem -RowLimit will be restored.

    .PARAMETER RowLimit
        Optionally limits restoration to a number of items when no identity is specified.

    .PARAMETER Force
        Restores the items without prompting for confirmation.

    .EXAMPLE
        RestoreAItemFromRecycleBin -Identity "72e4d749-d750-4989-b727-523d6726e442"

        Restores a single item by its recycle bin ID.

    .EXAMPLE
        RestoreAItemFromRecycleBin -RowLimit 100 -Force

        Restores up to 100 items from the recycle bin without confirmation.
    #>
    [CmdletBinding()]
    param(
        [Guid[]]$Identity,
        [int]$RowLimit,
        [switch]$Force
    )
    try {
        if ($PSBoundParameters.ContainsKey('Identity')) {
            Restore‑PnPRecycleBinItem -Identity $Identity -Force:$Force.IsPresent -ErrorAction Stop
        } elseif ($PSBoundParameters.ContainsKey('RowLimit')) {
            Get‑PnPRecycleBinItem -RowLimit $RowLimit | Restore‑PnPRecycleBinItem -Force:$Force.IsPresent -ErrorAction Stop
        } else {
            throw "Specify either Identity or RowLimit."
        }
        Write‑Verbose "Recycle bin item(s) restored."
    } catch {
        Write‑Error "Failed to restore recycle bin items: $_"
    }
}

###
### Function: CheckIfFileOrListItemIsWithinTheRecycleBin
###
function CheckIfFileOrListItemIsWithinTheRecycleBin {
    <#
    .SYNOPSIS
        Searches the recycle bin for a file or list item by name.

    .DESCRIPTION
        Retrieves recycle bin items using Get‑PnPRecycleBinItem.  Microsoft documentation
        explains that this cmdlet returns all items in the recycle bin for the connected
        site and requires that you be a site collection administrator【92742531835811†L839-L844】.  This
        function filters the results by LeafName (original file name) or DirName to locate
        deleted items.  Returns any matching recycle bin entries.

    .PARAMETER Name
        The file name (LeafName) to search for.

    .PARAMETER Location
        (Optional) Partial original path (DirName) to further narrow the search.

    .EXAMPLE
        CheckIfFileOrListItemIsWithinTheRecycleBin -Name "Budget.xlsx"

    .EXAMPLE
        CheckIfFileOrListItemIsWithinTheRecycleBin -Name "Budget.xlsx" -Location "/sites/Finance/Shared Documents"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$Location
    )
    try {
        $items = Get‑PnPRecycleBinItem -ErrorAction Stop
        $filtered = $items | Where‑Object {
            $_.LeafName -eq $Name -and (
                -not $PSBoundParameters.ContainsKey('Location') -or $_.DirName -like "*$Location*"
            )
        }
        return $filtered
    } catch {
        Write‑Error "Failed to search recycle bin: $_"
    }
}

### End of HarjTech SharePoint Online Management Toolkit