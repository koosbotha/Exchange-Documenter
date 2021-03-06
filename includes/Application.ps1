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
#region Writedoc
Function Write-Doc 
{
Param($body,$Head,$type)
Add-text "Starting Section $($Head)" 
$doc.Activate()
$selection=$word.Selection
$selection.EndOf(6) | Out-Null
$selection.TypeText("$([char]13)")
$selection.Style = $type
$selection.TypeText($Head)
$selection.TypeParagraph()
$selection.Style = "no spacing"
$selection.Font.Name="Segoe Pro"
$selection.Font.Size=10
$selection.TypeText($body)
$selection.TypeParagraph()
$selection.EndOf(6) | Out-Null
$selection.TypeParagraph()
}
#endregion Writedoc

#region Update Table
Function Update-Table
{
Param([Array]$TableContent,[int]$HeaderLignment,[int]$HeaderHeight)
$ErrorActionPreference = "SilentlyContinue"
$wdColor = 14121227
$TableRange = $doc.application.selection.range

$Columns = @($TableContent[0] | Get-Member -MemberType NoteProperty).count
$Rows = ($TableContent.Count)+1   #Add 1 for Column header 
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.AutoFitBehavior(2)
$table.Style = "Table Grid"
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1

$ColumnHeaders = $TableContent[0] | Get-Member -MemberType NoteProperty | select Name
$Cindex = 1 #Cell Index
Foreach ($ColumnH in $ColumnHeaders )
{
$xRow = 1
$trow = 2
$Table.Cell($xRow,$Cindex).Range.Orientation = $HeaderLignment
$Table.Cell($xRow,$Cindex).Shading.BackgroundPatternColor = $wdColor
$Table.Cell($xRow,$Cindex).Range.Font.Bold = $True
$Table.Cell($xRow,$Cindex).Range.Font.Color = "0000000"
$Table.Cell($xRow,$Cindex).Range.Text = $ColumnH.Name
If ($HeaderLignment -gt 0)
{
$Table.Rows.Item(1).Height = $HeaderHeight
}
    Foreach ($item in $TableContent )
    {

    $Table.Cell($tRow,$Cindex).Range.Text = $item.$($ColumnH.Name)
    ++$trow
    }
$Cindex++
}
#Autofit to Content
$table.AutoFitBehavior(1)

$selection.TypeParagraph()
$selection.EndOf(6) | Out-Null
$selection.TypeParagraph()

}
#endregion Update Table

#region Writelandingpage
Function Write-Landing 
{
    $doc.Activate()
    $selection=$word.Selection
    $selection.EndOf(6) | Out-Null
    $selection.Style = "Normal"
    $selection.Font.Size=10
    $selection.TypeText("$([char]13)")
    $selection.TypeText("$([char]13)")
    $selection.TypeText("$([char]13)")
    $selection.Style = "Normal"
    $selection.Font.Size=9
    $selection.ParagraphFormat.Alignment = 2
    $selection.TypeText("Prepared for$([char]13)")
    try{
    $selection.TypeText("$($ExOrg)$([char]13)")
    $selection.TypeText("$(Get-date -DisplayHint Date)$([char]13)")
    $selection.TypeText("Version 1.0$([char]13)")
    $selection.TypeText("Prepared by$([char]13)")
    $selection.TypeText("Author: $env:USERNAME$([char]13)" )
    $selection.TypeText("Email Address$([char]13)")
    $selection.TypeParagraph()
    $selection.TypeText("$([char]13)")
    $selection.ParagraphFormat.Alignment = 0
    $selection.EndOf(6) | Out-Null
    }
    Catch{
    $selection.TypeText("MyOrgName$([char]13)")
    $selection.TypeText("$(Get-date -DisplayHint Date)$([char]13)")
    $selection.TypeText("Version 1.0$([char]13)")
    $selection.TypeText("Prepared by$([char]13)")
    $selection.TypeText("Author: $env:USERNAME$([char]13)" )
    $selection.TypeText("Email Address$([char]13)")
    $selection.TypeParagraph()
    $selection.ParagraphFormat.Alignment = 0
    $selection.EndOf(6) | Out-Null
    }
    #Start writing Disclaimer information - Content extracted from txt in includes folder
    $selection.InsertNewPage()
    $selection.Font.Size=10
    $selection.Style = "Normal"
    $selection.Font.Name="Segoe Pro"
    $selection.TypeText([string](Get-Content -Path .\includes\Disclaimer.txt))
    $selection.TypeParagraph()
    $selection.EndOf(6) | Out-Null
    $selection.InsertNewPage()
    $selection.Font.Size=10
    $selection.Style = "Normal"
    $selection.Font.Name="Segoe Pro"
    $selection.TypeText("Table of Content")
    #TOC - Table of content Settings
    $tocrange = $doc.application.selection.range
    $useHeadingStyles = $true
    $upperHeadingLevel = 1
    $lowerHeadingLevel = 2
    $useFields = $false
    $tableID = "TOC1"
    $rightAlignPageNumbers = $true
    $includePageNumbers = $true
    $addedStyles = $null
    $useHyperlinks = $true
    $hidePageNumbersInWeb = $true
    $useOutlineLevels = $true
    $toc = $doc.TablesOfContents.Add($tocrange, $useHeadingStyles,$upperHeadingLevel, $lowerHeadingLevel, `
    $useFields, $tableID,$rightAlignPageNumbers, $includePageNumbers, $addedStyles,$useHyperlinks, $hidePageNumbersInWeb, $useOutlineLevels)
    $selection.InsertNewPage()
    $selection.EndOf(6) | Out-Null
}
#endregion Writelandingpage

