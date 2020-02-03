Add-PSSnapin FIMAutomation

$filename = "C:\Users\svreplante\SET_Description.html"

# $allSET = $null

if ($allSET -eq $null) {
    $allSET = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Set"
    $allGRP = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Group[MembershipLocked = true]"
}


function New-SetObject()
{
    param ($IdStr)

    $set = new-object PSObject
    $theSet = $allSet | where { $_.ResourceManagementObject.ResourceManagementAttributes.value -eq $IdStr }

    foreach ($attr in $theSet.ResourceManagementObject.ResourceManagementAttributes) {
        if ($attr.IsMultiValue -eq "True") {
            $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Values
        }
        else
        {
            if ($attr.AttributeName -eq "Filter") 
            {
                $t = [xml]($attr.Value)
                $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $t.Filter.InnerText
            }
            else
            {
                $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Value
            }
        }
    }

    return $set
}


function WriteToFile 
{
    PARAM($Str, $status)
    END 
    {
        if (($status -eq "begin") -and (Test-Path $filename)) {
            Remove-Item $filename
        }
    
        $Str | Out-file -FilePath $filename -Encoding utf8 -Append
    }
}

#Read All SETs
function GetAllSETInformation {

    $items = @()
    foreach ($set in $allSET) {

        $setAttributes = $set.ResourceManagementObject.ResourceManagementAttributes
   
        $theSET =   New-SetObject $set

        $item = new-object PSObject
        $item | add-member -type NoteProperty -Name "type" -Value "set"
        $item | add-member -type NoteProperty -Name "DN" -Value ($setAttributes | where { $_.AttributeName -eq "DisplayName" }).Value
        $item | add-member -type NoteProperty -Name "filter" -Value ($setAttributes | where { $_.AttributeName -eq "Filter" }).Value
        $item | add-member -type NoteProperty -Name "ExplicitMember" -Value ($setAttributes | where { $_.AttributeName -eq "ExplicitMember" }).Values.Count
        $item | add-member -type NoteProperty -Name "Description" -Value ($setAttributes | where { $_.AttributeName -eq "Description" }).Value
        $item | add-member -type NoteProperty -Name "grType" -Value "setType"

        $items += $item
    }

    return $Items

}

function GetAllGRPInformation {

    $items = @()
    foreach ($set in $allGRP) {

        $setAttributes = $set.ResourceManagementObject.ResourceManagementAttributes
        $type = ($setAttributes | where { $_.AttributeName -eq "Type" }).Value
   
        $theSET =   New-SetObject $set

        $item = new-object PSObject
        $item | add-member -type NoteProperty -Name "type" -Value "group"
        $item | add-member -type NoteProperty -Name "DN" -Value ($setAttributes | where { $_.AttributeName -eq "DisplayName" }).Value
        $item | add-member -type NoteProperty -Name "filter" -Value ($setAttributes | where { $_.AttributeName -eq "Filter" }).Value
        $item | add-member -type NoteProperty -Name "ExplicitMember" -Value ($setAttributes | where { $_.AttributeName -eq "ExplicitMember" }).Values.Count
        $item | add-member -type NoteProperty -Name "Description" -Value ($setAttributes | where { $_.AttributeName -eq "Description" }).Value
        $item | add-member -type NoteProperty -Name "grType" -Value $type

        $items += $item
    }

    return $Items

}

function HTMLCode {

    PARAM($part)
    END 
    {
        switch ($part) {
            "header" {
                return @"
<!DOCTYPE html>
<html>
<head>
<link href="SET_Description.css" rel="stylesheet">
</head>

<body>

<div class='title'><span class="titleGroups">CBG GROUPS</span> and <span class="titleSets">SETs</span> in MIM Portal</div>

<table>


"@ }
 
            "set" {   #0=DN, 1=Desc, 2=Filter, 3=NbrExplicite, 5=hiding Desc, 5=hiding Ext, 6=Odd True/False, 7=type, 8=error, 9=Group Type, 10=Group Type String
                return @"
<tr class="odd{6} {7}">
    <td class="{7}Name {9}"><div class="vertical">{10}</div></td>
    <td class="{7}Name">
        <div class=" {7}DN"><div  class="{7}DN">{0}</div></div>
        <div class=" {7}Description hide_{4}"><div  class=" {7}Desc">({1})</div></div>
        <div class=" {7}Explicite hide_{5} error_{8}">The set has {3} explicit user(s)</div>
        <div class=" {7}Filter">{2}</div>
    </td>
</tr>
"@ }

            "separator" { 
                return @"
<tr>
    <td class="separator" colspan=2></td>
</tr>
"@ }

            "sectionTitle" {  
                return @"
<div class='section {0}'><a id='{0}'>{1}</a></div>
"@ }
        }
    }
}

$items = GetAllSETInformation + GetAllGRPInformation
$items += GetAllGRPInformation
$items = $items| Sort-Object -Property type, DN
$nb = 0

#Header
WriteToFile (HTMLCode "header") "begin"
WriteToFile ""

#region Document all Transition in/out
foreach ($item in $items) {

    "Documenting item {0}" -f  $item.DN

    $grType = "Set"
    switch($item.grType) {
        "Distribution" { $grType = "DL" } 
        "MailEnabledSecurity" { $grType = "Security<br>Mail&nbsp;Enabled" } 
        "Security" { $grType = "Security" } 
        "setType" { $grType = "Set" } 
    }

    #0=DN, 1=Desc, 2=Filter, 3=NbrExplicite, 5=hiding Desc, 5=hiding Ext, 6=Odd True/False, 7=type
    $hidingExp =  ($item.ExplicitMember -eq 0)
    $hidingDesc =  ($item.Description -eq $null)
    $odd =  ($nb++ % 2) -eq 1
    $conflict = (($item.ExplicitMember -ne 0) -and ($item.filter -ne $null))
    WriteToFile ((HTMLCode "set") -f @($item.DN, $item.Description, $item.filter, $item.ExplicitMember, $hidingDesc, $hidingExp, $odd, $item.type, $conflict, $item.grType, $grType)) 

}
#endregion


WriteToFile "</table></body>"
