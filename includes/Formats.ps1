#############################################################################
#                                     			 		    				#
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
#                                     			 		    				#
#############################################################################

#region Formatrollup
Function Format-Rollup ([Array]$Content)
{
$returnValue = @()
$UServers = @($Content | Select ServerName -Unique)
$AllRollups = @($Content | Select DisplayName -Unique)
#add all possible RU for each object
$template =  New-Object PSObject
$template | Add-Member NoteProperty -Name " ServerName" -Value "" -Force
ForEach ($RU in $AllRollups)
            {
             $template | Add-Member NoteProperty -Name $RU.DisplayName -Value "" -Force
            }
$returnValue += $template 

ForEach ($item in $UServers)    
     {
        $obj = New-Object PSObject
        $Servers = $Content | ? {$_.ServerName -eq $item.ServerName}
        $obj | Add-Member NoteProperty -Name " ServerName" -Value $item.ServerName
        
        Foreach ($Server in $Servers)
            {
            $obj | Add-Member NoteProperty -Name $Server.DisplayName -Value "X" -Force
            } 
           
     $returnValue += $obj
     } 
 Return $returnValue 

}

#endregion Formatrollup

#region Format Array
Function Format-Array ([Array]$Content)
{
$ErrorActionPreference = "SilentlyContinue"
$returnValue = @()
$Objects = @($Content)

$template =  New-Object PSObject
$template | Add-Member NoteProperty -Name "Property" -Value "" -Force
ForEach ($O in $Objects)
{
$template | Add-Member NoteProperty -Name $O.Name -Value "" -Force
}
$returnValue += $template 


ForEach ($p in @($Content | Get-Member -MemberType NoteProperty))
{
$obj = New-Object PSObject
$obj | Add-Member NoteProperty -Name "Property" -Value $P.Name -Force
	ForEach ($O in $Objects)
	{
	$Value = $Objects | ?{$_.Name -eq $O.Name}
	$obj | Add-Member NoteProperty -Name $O.Name -Value (($Value).($p.Name) -join " ") -Force
	
	}
	$returnValue += $obj
}
 Return $returnValue 
}

#endregion Format Array