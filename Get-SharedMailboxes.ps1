<#
    .SYNOPSIS
    Apply-Permissions
   
    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, January 13th, 2021
    
    .DESCRIPTION
    This script assists in picking shared mailboxes to add to batches of users, based on the 
    permissions. Elegibility is determined by the users in a batch surpassing a percentage of permissions on 
    a shared mailbox. Optionally, a CSV can be specified containing mailboxes which should always be excluded.
    Also, a weightlist is used to give priority to permissions of certain types, e.g. usr permissions are set 
    to 2, group-based (grp) is set to 5. This information can also be used to control elegibility.
    
    Examples:
    - If UserA and UserB have permissions on SharedMailbox1, SharedMailbox1 can be considered to migrate together
      with mailboxes of both users. 
    - If UserA has permissions on SharedMailbox2 and UserB has not, a decisision needs to be made to include. The
      number of permissions for users part of the batch are set to 50%, which does not pass the default threshold
      of 75%.

    ❯ .\Get-SharedMailboxes.ps1  -UsersFile Users.csv -SharedMailboxFile Shared.csv -PermissionsFile Permissions.csv | Format-Table -a
    Batch D:\temp\Users.csv users: 3
    No mailboxes to exclude.
    Shared Mailboxes to consider: 1
    Permissions entries: 2 ..
    Constructing lookup table to speedup process..
    Constructing lookup table based on user..
    Processing..

    EmailAddress             TotalPerms TotalWeight InBatchPerms InBatchWeight Percentage IsExcluded Elegible
    ------------             ---------- ----------- ------------ ------------- ---------- ---------- --------
    shared@contoso.com                2           4            2             4        100      False     True

    Output:
    - EmailAddress  : E-mail address of the shared mailbox.
    - TotalPerms    : Total number of permissions applicable to this shared mailbox.
    - TotalWeight   : Total weight of all permissions on this shared mailbox.
    - InBatchPerms  : Total number of permissions of users specified in UsersFile applicable to this shared mailbox.
    - InBatchWeight : Total weight of all permissions of users specified in UsersFile on this shared mailbox.
    - Percentage    : Percentage of in-batch permissions.
    - IsExcluded    : Is the shared mailbox on the exclusion list.
    - Elegible      : Does percentage in-batch permissions exceed Treshold (default 75%), and is shared mailbox not excluded.
    
    .LINK
    http://eightwone.com
    
    .NOTES
    Requires that you are alread connected to Exchange or Exchange Online.

    Changelog
    --------------------------------------------------------------------------------
    0.1     Initial public release
    
    .PARAMETER UsersFile
    CSV file containing entries of mailboxes to process. Only column in the CSV file should be EmailAddress, containing 
    the e-mail address of mailboxes you want to process. This parameter is mandatory.

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
    you cannot specify PermissionsFile.

    .PARAMETER SharedMailboxFile
    CSV file containing entries of shared mailboxes to consider. Only column in the CSV file should be
    EmailAddress, containing the e-mail address of mailboxes you want to process. This parameter is mandatory.

    .PARAMETER ExcludeMailboxFile
    CSV file containing entries of mailboxes to exclude. Only column in the CSV file should be EmailAddress, containing the 
    e-mail address of mailboxes you want to process.

    .PARAMETER Treshold
    The treshold for the in-batch permissions for shared mailboxes to be considered elegible. Default is 75 (75%).

    .PARAMETER WeightList
    You can provide an hashtable to set the weight for various types of permissions. Default is @{ usr=2; all=1; grp=5; dom=1; guest=1; pub=1 }.
    Weight is not considered for shared mailboxes to be elegible, but can for example be used to decide tie-breakers.

    .EXAMPLE
    .\Get-SharedMailboxes.ps1  -UsersFile Users.csv -SharedMailboxFile Shared.csv -PermissionsFile Permissions.csv | Where-Object {$_.Elegible}
    
    Show shared mailboxes from Shared.csv which are elegible for migration, and can be added to batch Users.csv based on permissions from Permissions.csv
#>
#Requires -Version 3.0
[cmdletbinding( SupportsShouldProcess=$true, DefaultParameterSetName='General')]
param(
    [parameter( Mandatory=$true, ParameterSetName='FileMode')]
    [parameter( Mandatory=$true, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$UsersFile,
    [parameter( Mandatory=$true, ParameterSetName='FileMode')]
    [parameter( Mandatory=$true, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$SharedMailboxFile,
    [parameter( Mandatory=$true, ParameterSetName='FileMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]    
    [string]$PermissionsFile,
    [parameter( Mandatory=$true, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Container})]
    [string]$PermissionsFolder,  
    [parameter( Mandatory=$false, ParameterSetName='FolderMode')]
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [int]$Treshold= 75,
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [parameter( Mandatory=$false, ParameterSetName='FolderMode')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [string]$ExcludeMailboxFile,
    [parameter( Mandatory=$false, ParameterSetName='FileMode')]
    [parameter( Mandatory=$false, ParameterSetName='FolderMode')]
    [hashtable]$WeightList= @{ usr=2; all=1; grp=5; dom=1; guest=1; pub=1 }
)

$Users=  Import-csv $UsersFile -Header 'EmailAddress'
Write-Host ('Batch {0} users: {1}' -f $UsersFile, ($Users | Measure-Object).Count)

$ExcludeMailbox= @{}
If( $ExcludeMailboxFile) {
    Import-Csv -Path $ExcludeMailboxFile | ForEach-Object { $ExcludeMailbox[ $_.EmailAddress]= $True }
    Write-Host ('Mailboxes to exclude: {0}' -f ($ExcludeMailbox | Measure-Object).Count)
}
Else {
    Write-Host ('No mailboxes to exclude.')
}

$SharedMbx= @{}
Import-Csv -Path $SharedMailboxFile| ForEach-Object { $SharedMbx[ $_.EmailAddress]= $True }
Write-Host ('Shared Mailboxes to consider: {0}' -f ($SharedMbx | Measure-Object).Count)

$Perms= Import-Csv -Path $PermissionsFile -Delimiter ';' -Header 'Mailbox','Path','Type','Delegate','Perms','PermsDetails'
Write-Host ('Permissions entries: {0} ..' -f ($Perms | Measure-Object).Count)

$LookupPerms= @{}
Write-Host 'Constructing lookup table to speedup process..'
$GroupPerMailbox= $perms | Group-Object Mailbox
ForEach( $Entry in $GroupPerMailbox) {
    If( $LookupPerms[ $Entry.Name]) {
	$LookupPerms[ $Entry.Name]+= $Entry.Group
    }
    Else {
        $LookupPerms[ $Entry.Name]= $Entry.Group
    }    
}
Write-Host 'Constructing lookup table based on user..'
$GroupPerUser= $perms | Group-Object Delegate
ForEach( $Entry in $GroupPerUser) {
    If( $LookupPerms[ $Entry.Name]) {
	$LookupPerms[ $Entry.Name]+= $Entry.Group
    }
    Else {
        $LookupPerms[ $Entry.Name]= $Entry.Group
    }    
}

Write-Host 'Processing..'

ForEach( $Mbx in $SharedMbx.getEnumerator()) {

    $PermsTotal= $LookupPerms[ $Mbx.Name] 
    $TotalCount= ($PermsTotal | Measure-Object).Count 

    $BatchCount= 0
    $TotalWeight= 0
    $InBatchWeight= 0
    ForEach( $Perm in $PermsTotal) {

        $Weight= $WeightList[ $Perm.Perms]+0
        $TotalWeight+= $Weight

        ForEach( $User in $Users) {
            If( $Perm.Mailbox -ieq $User.EmailAddress -or $Perm.Delegate -ieq $User.EmailAddress) { 
                $BatchCount++
                $InBatchWeight+= $Weight
            }
        }
    }

    If( $TotalCount -gt 0) {
        $Pct= [int]( $BatchCount/ $TotalCount * 100) 
    }
    Else {
        $Pct= 0
    }

    $IsExcluded= $null -ne $ExcludeMailbox[ $Mbx.Name]

    $obj= [pscustomobject]@{
        EmailAddress= $Mbx.Name
        TotalPerms= $TotalCount
        TotalWeight= $TotalWeight
        InBatchPerms= $BatchCount
        InBatchWeight= $InBatchWeight
        Percentage= $Pct
        IsExcluded= $IsExcluded
        Elegible= ($Pct -ge $Treshold) -and (-not $IsExcluded) 
    }
    $obj
}