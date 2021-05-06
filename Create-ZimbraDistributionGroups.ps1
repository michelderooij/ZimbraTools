param(
    [Parameter( Mandatory=$true)]
    [string]$DGFile,
    [Parameter( Mandatory=$false)]
    [string]$Filter='*',
    [Parameter( Mandatory=$false)]
    [switch]$Create,
    [Parameter( Mandatory= $false)]
    [string]$DefaultManager,
    [Parameter( Mandatory= $false)]
    [string]$DGPrefix
)

If( !( Test-Path $DGFile)) {
    Write-Error 'DG file not found or specified'
}
$DGFileData=  Get-Content $DGFile | Where {$_ -notlike '#*'}
Write-Host ('Reading {0}' -f $DGFile)

$DGData= @{}
$ACEStack= [System.Collections.ArrayList]@()
$validRecipients= @{}
$ActiveDG= $null
$MembersMode= $false
$primaryMode= $true

ForEach($Line in $DGFileData) {

# For Testing:
#    $Line= $Line -replace '@zzv.nl', '@myexchangelabs.com'
#    $Line= $Line -replace '@zzvnl.onmicrosoft.com', '@myexchangelabs.onmicrosoft.com'
#    $Line= $Line -replace '@zzvnl.mail.onmicrosoft.com', '@myexchangelabs.mail.onmicrosoft.com'
#    $Line= $Line -replace '@warmande.nl', '@hylabo.com'

    Switch -RegEx($Line) {
        '^zimbraMailAlias:.*$' {
            If( $primaryMode) {
                $Name= ($Line -split ': ')[1]
                Write-Verbose ('Parsing {0}' -f $Name)
                $DGData[ $Name]= [pscustomobject]@{ 
                    Name= '{0}{1}' -f $DGPrefix, ($Name -split '@')[0]
                    Identity= $Name
                    EmailAddresses= [System.Collections.ArrayList]@()
                    Members= [System.Collections.ArrayList]@()
                    Status= $null
                    ManagedBy= [System.Collections.ArrayList]@()
                    AuthSenders= [System.Collections.ArrayList]@()
                    ChildrenDG= [System.Collections.ArrayList]@()
                }
                $null= $DGData[ $Name].EmailAddresses.Add( 'SMTP:{0}' -f $Name)
                If( $ACEStack) {
                    # Note: ^=Add, -=Remove, Nothing=Just set?
                    ForEach( $ACE in $ACEStack) {
                        $Item= $ACE -split ' ' 
                        Switch( $Item[2]) {
                            'ownDistList' {
                                Write-Verbose ('{0}: Processing Owner {1} ({2})' -f $Name, $Item[0], $Item[1])
                                $null= $DGData[ $Name].ManagedBy.Add( $Item[0])
                            }
                            '^sendToDistList' {
                                Write-Verbose ('{0}: Processing Allowed Sender {1} ({2})' -f $Name, $Item[0], $Item[1])
                                $null= $DGData[ $Name].AuthSenders.Add( @($Item[0], $Item[1]))
                            }
                            'sendToDistList' {
                                Write-Verbose ('{0}: Processing Allowed Sender {1} ({2})' -f $Name, $Item[0], $Item[1])
                                $null= $DGData[ $Name].AuthSenders.Add( @($Item[0], $Item[1]))
                            }
                            '-modifyDistributionList' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-removeDistributionListMember' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-sendToDistList' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-renameDistributionList' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-deleteDistributionList' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-addDistributionListAlias' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-removeDistributionListAlias' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            '-addDistributionListMember' {
                                Write-Verbose ('{0}: Ignoring ACE {1}' -f $Name, $ACE)
                            }
                            default {
                                Write-Warning ('{0}: Unknown ACE: {1}' -f $Name, $ACE)
                            }
                        }
                    }
                    $ACEStack= [System.Collections.ArrayList]@()
                }
                $activeDG= $Name
                $MembersMode= $false
                $primaryMode= $false
            }
            Else {
                Write-Verbose ('Adding alias {0} to DG {1}' -f ($Line -split ': ')[1], $activeDG)
                $null= $DGData[ $activeDG].EmailAddresses.Add( 'smtp:{0}' -f ($Line -split ': ')[1])
            }
        }            
        '^zimbraMailStatus:.*$' {
            $DGData[ $activeDG].Status= ($Line -split ': ')[1]
            $MembersMode= $false
            $primaryMode= $true
        }
        'zimbraACE:.*$' {
            $null= $ACEStack.Add( ($Line -split ': ')[1])
            $MembersMode= $false
            $primaryMode= $true
        }
        '^members$' {
            $MembersMode= $true    
            $primaryMode= $true
        }
        default {
            $primaryMode= $true
            If( $MembersMode) {
                Write-Verbose ('{0}: Member {1}' -f $activeDG, $Line)
                $null= $DGData[ $activeDG].Members.Add( $Line)
            }
            Else {
                If( [string]::IsNullOrEmpty( $Line)) {
                    # Blank line
                }
                Else {
                    Write-Host ('End of file')
                }
            }
        }
    }
}

# Keep track of DGs that are member of another DG (so we can process them in proper order)
ForEach( $DG in $DGData.GetEnumerator()) {
    ForEach( $Member in $DGData[ $DG.Name].Members) {
        If( $DGData[ $Member]) {
            $null= $DGData[ $Member].ChildrenDG.Add( $DG.Name)
        }
    }
}

$DGData= $DGData.GetEnumerator() | Select -expandProperty Value

# Determine data set to work with
$DGProcessed= $DGData | Sort -Property @{Expression= {$_.ChildrenDG.Count}; Descending= $True} | Where {$_.Name -like $Filter}

If( $Create) {
    ForEach( $DG in $DGProcessed) {
        Write-Host ('Processing {0} ({1})' -f $DG.name, $DG.Identity)
        $GroupObj= Get-DistributionGroup -Identity $DG.Identity -ErrorAction SilentlyContinue
        If( $GroupObj) {
            $Identity= $GroupObj.Name
            Write-Host ('Updating {0} ..' -f $Identity)
        }
        Else {
            $CreateParms= @{
                Name= $DG.Name
                Type= 'Distribution'
                IgnoreNamingPolicy= $true
            }
            # Setting primarySmtpAddress; setting it with splatting doesn't work
            $GroupObj= New-DistributionGroup -primarySmtpAddress $DG.Identity @CreateParms
            $Identity= $NewGroup.Name
            Write-Host ('Created {0} ..' -f $Identity)
        }        
        $UpdateParms= @{
            Name= $DG.Name
            EmailAddresses= $DG.EmailAddresses
            AcceptMessagesOnlyFromSendersOrMembers= $null
        }
        If( $DG.ManagedBy) {
            $UpdateParms.ManagedBy= $DG.ManagedBy | Get-EXORecipient -ErrorAction SilentlyContinue -Verbose:$False | Select -expandProperty primarySmtpAddress
            Write-Host ('Validated {1} of {0} owners for {2}..' -f ($DG.ManagedBy | measure).Count, ( $UpdateParms.ManagedBy | Measure).Count, $DG.Name)
        }
        If( $DG.AuthSenders) {
            $UpdateParms.RequireSenderAuthenticationEnabled= $true
            $UpdateParms.AcceptMessagesOnlyFromSendersOrMembers= [System.Collections.ArrayList]@()
            ForEach( $Sender in $DG.AuthSenders) {
                Switch( $Sender[1]) {
                    'all' {
                        # Only internal senders (=RequireSenderAuthenticationEnabled)                                       
                    }
                    'usr' {
                        If( Get-EXORecipient -Identity $Sender[0] -ErrorAction SilentlyContinue -Verbose:$False) {
                            $null= $UpdateParms.AcceptMessagesOnlyFromSendersOrMembers.Add( $Sender[0])
                        }
                    }
                    'grp' {
                        If( Get-EXORecipient -Identity $Sender[0] -ErrorAction SilentlyContinue -Verbose:$False) {
                            $null= $UpdateParms.AcceptMessagesOnlyFromSendersOrMembers.Add( $Sender[0])
                        }
                    }
                    default {
                        Write-Warning ('{0}: Unknown ACE AuthSender directive: {1}' -f $DG.Name, $Sender[1]) 
                    }
                }
            }
        }
        Else {
            # When nothing is specified, everyone is allowed to send to DG, including external
            $UpdateParms.AcceptMessagesOnlyFromSendersOrMembers= $null
        }

        If( $DG.Status -eq 'disabled') {
            # When disabled, mimic behavior by only allowing authenticated self sender
            $UpdateParms.RequireSenderAuthenticationEnabled= $true
            $UpdateParms.AcceptMessagesOnlyFromSendersOrMembers= $DG.Identity
        }
        If(!( $UpdateParms.ManagedBy)) {
            If(!( $DefaultManager)) {
                $DefaultManager= (Get-PSSession).Where({$_.ComputerName -like '*.office365.com' -and $_.State -eq 'Opened'})[0].RunSpace.ConnectionInfo.Credential.UserName
            }
            Write-Host ('{0} has no owner; appointing me ({1})' -f $DG.Name, $DefaultManager)
            $UpdateParms.ManagedBy= $DefaultManager
            $UpdateParms.BypassSecurityGroupManagerCheck= $true
        }

        Set-DistributionGroup -Identity $DG.Identity @UpdateParms

        If( $DG.Members) {
            # First, filter the unique ID, later filter unique again to handle specified aliases instead of primary SMTP addresses
            $ValidMembers= $DG.Members | Select -Unique | Get-EXORecipient -ErrorAction SilentlyContinue -Verbose:$False | Select -Unique -expandProperty primarySmtpAddress 
            Write-Host ('Added {1} validated members of {0} for {2}..' -f ($DG.members | measure).Count, ( $ValidMembers | Measure).Count, $DG.Name)
            Update-DistributionGroupMember -Identity $DG.Identity -Members $ValidMembers -BypassSecurityGroupManagerCheck -Confirm:$false
        }
    }
}
Else {
    # Report only
    $DGProcessed
}