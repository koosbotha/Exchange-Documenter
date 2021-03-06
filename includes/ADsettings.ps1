#############################################################################
#                                     			 		    #
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 		    #
#############################################################################
Import-Module ActiveDirectory 

Function Get-FSMO
{
$ErrorActionPreference = "SilentlyContinue"

$Result = @()
$Domains = @((Get-ADForest).domains)
Foreach ($DOM in $Domains)
    {
    $ADForest = Get-ADForest | Select DomainNamingMaster,SchemaMaster
    $Domain = Get-ADDomain $DOM
    $obj = New-Object PSObject
    $obj | Add-Member NoteProperty -Name "DNSRoot"              -Value $Domain.DNSRoot -Force
    $obj | Add-Member NoteProperty -Name "PDCEmulator"          -Value $Domain.PDCEmulator -Force
    $obj | Add-Member NoteProperty -Name "RIDMaster"            -Value $Domain.RIDMaster -Force
    $obj | Add-Member NoteProperty -Name "InfrastructureMaster" -Value $Domain.InfrastructureMaster -Force
    $obj | Add-Member NoteProperty -Name "DomainNamingMaster"   -Value $ADForest.DomainNamingMaster -Force
    $obj | Add-Member NoteProperty -Name "SchemaMaster"         -Value $ADForest.SchemaMaster -Force
    $Result += $obj
    }
    Return $Result
}

Function get-Sites
{
$ErrorActionPreference = "Stop"
Try{
#Determine Site With Exchange Servers
$ExInstalledSites = @((Get-ExchangeServer | ?{  $ExcludedServers -notcontains $_.Name} | select Site -Unique).site.rdn.EscapedName)
$ExchangeSites = @(Get-ADSite | ?{$ExInstalledSites -contains $_.Name } | select Name,HubSiteEnabled,ExchangeVersion )
Return $ExchangeSites
}
Catch {
Add-text $_.Message 
      }
}

Function get-SiteLinks
{
$ErrorActionPreference = "Stop"
try{
$Return = @()
$ExInstalledSites = @(Get-ExchangeServer | ?{$ExcludedServers -notcontains  $_.Name}| select Site -Unique)
$Allsites = @(Get-AdSiteLink  | select Cost,Sites,ADCost,ExchangeCost,MaxMessageSize,ExchangeVersion,Name)
ForEach($ExSite in $ExInstalledSites)
    {
            ForEach ($I in $Allsites)
                    {
                    #Create a new object for SiteNames included in SiteLink
                    $obj = New-Object PSObject
                    $obj | Add-Member NoteProperty -Name "Cost"           -Value $i.Cost -Force
                    $obj | Add-Member NoteProperty -Name "ADCost"         -Value $i.ADCost -Force
                    $obj | Add-Member NoteProperty -Name "ExchangeCost"   -Value $i.ExchangeCost -Force
                    $obj | Add-Member NoteProperty -Name "MaxMessageSize" -Value $i.MaxMessageSize -Force
                    $obj | Add-Member NoteProperty -Name "Name"           -Value $i.Name -Force
                    $inclSites = @()

                    ForEach ($MSite in @($I.Sites))
                        {
                    #Check if the Site Is part of any Exchange Server Related Sites
                            IF ($ExSite.Site.rdn.EscapedName -eq $MSIte.rdn.EscapedName)
                            {                        
                            ForEach ($MSite in @($I.Sites)){$inclSites += $MSIte.rdn.EscapedName}
                            }
                        } 
                        If ($inclSites.count -gt 0) 
                        {
                        [String] $Sitevalue = ""
                        Foreach ($s in $inclSites){
                        $Sitevalue = $Sitevalue + $s + " "
                        }
                        $obj | Add-Member NoteProperty -Name "Site" -Value $Sitevalue -Force
                        $Return += $obj
                        } Else {$obj = ""}  
                    }
    }
   Return $Return
   }
   Catch {
   Add-text $_.Message #-ForegroundColor Red
   }
}