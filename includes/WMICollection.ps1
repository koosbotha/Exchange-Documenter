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

Function Get-WMI($class,$props)
{
$EnvServers = Get-EXServerNodes
    $Result = @()
    Foreach ($server in $EnvServers)
    {
    $ErrorActionPreference = "Stop"
    # Check if server is available
    $ping = New-Object –TypeName System.Net.Networkinformation.Ping
    Try { 
        $status = ($ping.Send($($server.Name))).Status 
        #Status Check returned with true
        }
    Catch {
        $status = "Failure"
        Add-text "Cannot reach server $($server.Name) . It will be excluded from results." #-ForegroundColor red
        }
    if ($? -eq "True" -and $status -eq "Success")
        {
            If ($status -eq "Success")
                {
				 
                $ErrorActionPreference = "SilentlyContinue"
                #Add-text "Collecting $($class.replace('win32_','')) Information for $($server.Name)" 
				$objItems = @(Get-WmiObject -Computer $server.Name -Namespace 'root\cimv2' -Query $class -ErrorAction SilentlyContinue  | Select $props)
                
                            Foreach ($Objitem in $objItems)
                            {
                                $obj = New-Object PSObject
                                $obj | Add-Member NoteProperty -Name "Server" -Value $server.Name
                                    ForEach ($NoteProperty in $props)
                                    {
                                    $obj | Add-Member NoteProperty -Name $NoteProperty -Value ($objItem.$NoteProperty -join " ")
                                    }
                                    $Result += $obj
                            }

                }
          }
       }
   Return $Result
}