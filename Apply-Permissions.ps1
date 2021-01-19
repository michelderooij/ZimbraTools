<#
    .SYNOPSIS
    Apply-Permissions
   
    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, January 19th, 2021
    
    .DESCRIPTION
    This script allows you map exported Zimbra mailbox permissions to Exchange Online, where
    possible. That is, where possible, as some permissions are not realizable using Exchange which, such
    as read-only mailboxes.

    When performing the mapping, the following folders are translated to their (localized) well-known folders:
    - /Inbox        
    - /Calendar     
    - /Tasks
    - /Sent

    Regarding translating Zimbra permissions to Exchange permissions, the following mappings are made:
    
    Perms   Type                   PermsDetails        Exchange Permission      Level
    ==========================================================================================
    <empty> *                      <empty>             None                     Folder
    usr     *                      r                   Reviewer                 Folder
    usr     *                      rwidx               Editor                   Folder
    usr     *                      rwidxa/rwidxp       Owner                    Folder
    usr/grp doc/tsk/msg/cont/apt   r                   Reviewer                 Folder
    usr/grp doc/tsk/msg/cont/apt   rwidx               Editor                   Folder
    usr/grp doc/tsk/msg/cont/apt   rwidxa              Owner                    Folder
    usr     ROOT                   rwidxa/rwidx        FullAccess+SendAs        Mailbox
    usr     ROOT                   r                   ReadPermission           Mailbox
    pub     appointment            r                   Default/LimitedDetails   Folder (Calendar)
    guest   appointment            r                   Anonymous/AvailOnly      Folder (Calendar)
    all     contact                r                   Default/Reviewer         Folder
    dom     contact                r                   Default/Reviewer         Folder
    grp     ROOT                   rwidxa              FullAccess+SendAs        Mailbox

    For reference, these are the possible Zimbra permissions:
    - (r)ead - search, view overviews and items
    - (w)rite - edit drafts/contacts/notes, set flags
    - (x) action - workflow actions, like accepting appoitnments
    - (i)nsert - copy/add to directory, create subfolders
    - (d)elete - delete items and subfolders, set Deleted flag
    - (a)dminister - delegate admin and change permissions 

    The script supports WhatIf for reporting and evaluation, and requires Confirm for non-interactive processing of permissions.
    
    Sample output:
    --------------
    > .\Apply-Permissions.ps1 -UsersFile Users.csv -PermissionsFile .\Permissions.csv | Format-Table -a
    Batch Users.csv users: 2
    Total Permissions entries: 2 ..
    In-Batch Permissions entries: 2 ..

    Mailbox             Path   Type    Perms       ExFolderId                          Action
    -------             ----   ----    -----       ----------                          ------
    shared@contoso.com  /      ROOT    usr[rwidxa] postmaster@contoso.com:\            Grant FullAccess+SendAs on shared@contoso.com to philip@consoso.com
    shared@contoso.com  /Inbox message usr[rwidxa] postmaster@contoso.com:\Postvak In  Grant Owner on shared@contoso.com:\Postvak In to francis@contoso.com

    .LINK
    http://eightwone.com
    
    .NOTES
    Requires that you are alread connected to Exchange or Exchange Online.

    Changelog
    --------------------------------------------------------------------------------
    1.0    Initial public release

    .PARAMETER UsersFile
    CSV file containing entries of mailboxes to process. Only column in the CSV file should be
    EmailAddress, containing the e-mail addresses of mailboxes you want to process. This parameter
    is mandatory.

    .PARAMETER PermissionsFile
    Specifies the CSV file containing the exporter Zimbra mailbox permissions. File does not need a header, but should
    contain the following elements per line, using ';' as seperator:
    info@contoso.com;/;ROOT;p.mortimer@contoso.com;usr;rwidxa

    Elements:
    - info@contoso.com          E-mail address of Mailbox
    - /                         Path
    - ROOT                      Item type or ROOT for mailbox-level (Top of Information Store)
    - p.mortimer@contoso.com    Delegate
    - usr                       Type of assignment
    - rwidxa                    Permissions

    Specifying PermissionsFile or PermissionsFolder is mandatory. When specifying PermissionsFile, you cannot specify PermissionsFolder.

    .PARAMETER PermissionsFolder
    Specifies the folder containing individual files with per-mailbox permissions. The layout of these files need to be
    the same as  when a PermissionsFile. Specifying PermissionsFile or PermissionsFolder is mandatory. When specifying PermissionsFolder,
    you cannot specify PermissionsFile. Note that permission files have no extension and similar to PermissionsFile, no header.

    .PARAMETER AutoMapping
    Indicates if AutoMapping should be used when assigning Full Access permissions. Default is false.

    .PARAMETER SendNotificationToUser
    Whether sharing invitations are sent to delegates when appropriate user-assigned calender permissions are applied, 
    i.e. AvailabilityOnly, LimitedDetails, Reviewer or Editor.

    .PARAMETER SharingPermissionFlags
    Whether calendar delegate permissions are configured when appropriate user-assigned calendar permissions are applied, i.e. Editor.
    You can specify one or more of values None (Default), Delegate or CanViewPrivateItems. For more information, see
    https://docs.microsoft.com/en-us/powershell/module/exchange/add-mailboxfolderpermission.

    .EXAMPLE
    Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFile Permissions.csv

    Applies permissions from Permissions.csv using mailboxes specified in Users.csv.

    .EXAMPLE
    Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFolder C:\PermissionFiles -Confirm:$false 

    Applies permissions using permission files located in C:\PermissionFiles, using mailboxes specified in Users.csv. Script will not ask
    for confirmation for each operation.

    .EXAMPLE
    Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFile Permissions.csv -WhatIf:$true | Export-Csv PermissionReport.csv -NoTypeInformation    

    Export what permissions would be applied from Permissions.csv using mailboxes specified in Users.csv. Output is exported to PermissionReport.csv.
#>
#Requires -Version 3.0
[cmdletbinding( SupportsShouldProcess=$true, DefaultParameterSetName='General')]
param(
    [parameter( Mandatory=$true, ParameterSetName='FileMode')]
    [parameter( Mandatory=$true, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$UsersFile,
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [parameter( Mandatory=$false,ParameterSetName='FolderMode')]
    [bool]$AutoMapping=$false,
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [parameter( Mandatory=$false,ParameterSetName='FolderMode')]
    [ValidateSet('None', 'Delegate', 'CanViewPrivateItems')]
    [string[]]$SharingPermissionFlags='None',
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [parameter( Mandatory=$false,ParameterSetName='FolderMode')]
    [bool]$SendNotificationToUser=$false,
    [parameter( Mandatory=$true, ParameterSetName='FileMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$PermissionsFile,
    [parameter( Mandatory=$true, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Container})]
    [string]$PermissionsFolder
)

If( !( Get-Command Get-MailboxFolderStatistics -ErrorAction SilentlyContinue)) {
    Throw 'Not connected to Exchange Online'
}

# Current Exchange session is Exchange Online
$local:EXOMode= (Get-OrganizationConfig).Name -ilike '*.onmicrosoft.com'

$Users=  Import-csv -Path $UsersFile -Header 'EmailAddress' 
Write-Host ('Batch {0} users: {1}' -f $UsersFile, ($Users | Measure-Object).Count)

$LookupUsers= @{}
$Users | ForEach-Object { $LookupUsers[ $_.EmailAddress]= $_ }

If( $PSCmdlet.ParameterSetName -eq 'FileMode') {
    $Perms= Import-Csv -Path $PermissionsFile -Delimiter ';' -Header 'Mailbox','Path','Type','Delegate','Perms','PermsDetails'
    Write-Host ('Total Permissions entries: {0} ..' -f ($Perms | Measure-Object).Count)
}
Else {
    Get-ChildItem -Path $PermissionsFolder -Include *.*
}

$BatchPerms= $Perms | Where-Object { $LookupUsers[ $_.Mailbox] -or ($_.Delegate -and $LookupUsers[ $_.Delegate]) } 
Write-Host ('Applicable permissions entries: {0} ..' -f ($BatchPerms | Measure-Object).Count)

ForEach( $Perm in $BatchPerms) {

    If(!( Get-Mailbox -Identity $Perm.Mailbox -ErrorAction SilentlyContinue)) {

        Write-Host ('Mailbox not found used in permission: {0}' -f $Perm.Mailbox) -ForegroundColor Red
    }
    Else {

        $obj=[pscustomobject]@{
            Mailbox= $Perm.Mailbox
            Path= $Perm.Path
            Type= $Perm.Type
            Perms= '{0}[{1}]' -f $Perm.Perms, $Perm.PermsDetails
            Notify= $null
            Flags= $null
        }

        # Lookup the actual folder when processing certain 'well-known' folders to accomodate for localization.
        # Note: We Where-filter on Foldertype 'again', as for example FolderType Calendar will also return unmoveable folder 'Birthdays'.
        $IsCal= $false
        Switch -regex( $Perm.Path) {
            '^/Inbox(/.*)?$' {
                $FolderName = (Get-MailboxFolderStatistics -Identity $Perm.Mailbox -FolderScope Inbox | Where-Object {-not $_.Movable -and $_.FolderType -eq 'Inbox'} | Select-Object -First 1).Name     
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, (($Perm.Path -replace '/Inbox', ('\{0}' -f $FolderName)) -replace '/', '\')
            }
            '^/Calendar(/.*)?$' {
                $FolderName = (Get-MailboxFolderStatistics -Identity $Perm.Mailbox -FolderScope Calendar | Where-Object {-not $_.Movable -and $_.FolderType -eq 'Calendar'} | Select-Object -First 1).Name 
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, (($Perm.Path -replace '/Calendar', ('\{0}' -f $FolderName)) -replace '/', '\')
                $IsCal= $true
            }
            '^/Tasks(/.*)?$' {
                $FolderName = (Get-MailboxFolderStatistics -Identity $Perm.Mailbox -FolderScope Tasks | Where-Object {-not $_.Movable -and $_.FolderType -eq 'Tasks'} | Select-Object -First 1).Name 
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, (($Perm.Path -replace '/Tasks', ('\{0}' -f $FolderName)) -replace '/', '\')
            }
            '^/Contacts(/.*)?$' {
                $FolderName = (Get-MailboxFolderStatistics -Identity $Perm.Mailbox -FolderScope Contacts | Where-Object {-not $_.Movable -and $_.FolderType -eq 'Contacts'} | Select-Object -First 1).Name 
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, (($Perm.Path -replace '/Contacts', ('\{0}' -f $FolderName)) -replace '/', '\')
            }
            '^/Sent(/.*)?$' {
                $FolderName = (Get-MailboxFolderStatistics -Identity $Perm.Mailbox -FolderScope SentItems | Where-Object {-not $_.Movable -and $_.FolderType -eq 'SentItems'} | Select-Object -First 1).Name 
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, (($Perm.Path -replace '/Sent', ('\{0}' -f $FolderName)) -replace '/', '\')
            }
            default {
                $ExchangeFolderId= '{0}:{1}' -f $Perm.Mailbox, ($Perm.Path -replace '/', '\')
            }
        }

        # Add the effective Exchange folder to the output object
        $obj | Add-Member -Type NoteProperty -Name 'ExFolderId' -Value $ExchangeFolderId

        # Determine course of action
        Switch( $Perm.Type) {
            'ROOT' {
                Switch -regex( $Perm.PermsDetails) {
                    '^rwidx(a)?$' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant FullAccess+SendAs on {0} to {1}' -f $Perm.Mailbox, $Perm.Delegate)
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxPermission -Identity $Perm.Mailbox -User $Perm.Delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping | Out-Null
                            If( $local:EXOMode) {
                                Add-RecipientPermission -Identity $Perm.Mailbox -Trustee $Perm.Delegate -AccessRights SendAs -Confirm:$False | Out-Null
                            }
                            Else {
                                Add-ADPermission -Identity $Perm.Mailbox -User $Perm.Delegate -AccessRights ExtendedRight -ExtendedRights 'Send As' -Confirm:$False | Out-Null
                            }
                        }
                    }
                    '^r$' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights ReadPermission -InheritanceType All | Out-Null
                        }
                    }
                    default {
                        Write-Host ('Unsupported permission combination: {0}-{1}-{2}[{3}]' -f $Perm.Mailbox, $Perm.Type, $Perm.Perms, $Perm.PermsDetails) -ForegroundColor Red
                    }
                }
            }
            default {
                switch( $Perm.Perms) {
                    $null {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant None on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights None | Out-Null
                        }
                    }
                    'pub' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, 'Default')
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxFolderPermission -Identity $ExchangeFolderId -User Default -AccessRights LimitedDetails | Out-Null
                        }
                    }
                    'guest' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, 'Anonymous')
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxFolderPermission -Identity $ExchangeFolderId -User Anonymous -AccessRights AvailabilityOnly | Out-Null
                        }
                    }
                    'all' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, 'Default')
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxFolderPermission -Identity $ExchangeFolderId -User Default -AccessRights Reviewer | Out-Null
                        }
                    }
                    'dom' {
                        $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, 'Default')
                        If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                            Add-MailboxFolderPermission -Identity $ExchangeFolderId -User Default -AccessRights Reviewer | Out-Null
                        }
                    }
                    default {
                        # usr/grp
                        switch( $Perm.PermsDetails) {
                            'r' {
                                $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Reviewer on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                                If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                                    If( $IsCal) {
                                        Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Reviewer -SharingPermissionFlags $SharingPermissionFlags -SendNotificationToUser $SendNotificationToUser | Out-Null
                                        $obj.Notify= $SendNotificationToUser
                                        $obj.Flags= $SharingPermissionFlags
                                    }
                                    Else {
                                        Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Reviewer | Out-Null
                                    }
                                }
                            }
                            'rwidx' {
                                $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Editor on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                                If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                                    If( $IsCal) {
                                        Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Editor -SharingPermissionFlags $SharingPermissionFlags -SendNotificationToUser $SendNotificationToUser | Out-Null
                                        $obj.Notify= $SendNotificationToUser
                                        $obj.Flags= $SharingPermissionFlags
                                    }
                                    Else {
                                        Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Editor | Out-Null
                                    }
                                }
                            }
                            'rwidxa' {
                                $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Owner on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                                If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                                    Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Owner | Out-Null
                                }
                            }
                            'rwidxp' {
                                $obj | Add-Member -Type NoteProperty -Name 'Action' -Value ('Grant Owner on {0} to {1}' -f $ExchangeFolderId, $Perm.Delegate)
                                If ($PSCmdlet.ShouldProcess($obj.Action, 'Are you sure you want to perform this action?')) {  
                                    Add-MailboxFolderPermission -Identity $ExchangeFolderId -User $Perm.Delegate -AccessRights Owner | Out-Null
                                }
                            }
                            default {
                                Write-Host ('Unsupported permission combination: {0}-{1}-{2}[{3}]' -f $Perm.Mailbox, $Perm.Type, $Perm.Perms, $Perm.PermsDetails) -ForegroundColor Red
                            }
                        }
                    }
                }
            }
        }
    }
    $obj
}
