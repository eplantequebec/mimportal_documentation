#Voir https://github.com/eplantequebec/mimportal_documentation

Add-PSSnapin FIMAutomation

$filename = "MPR_Description.html"

# $AllMPR = $null

if ($allMPR -eq $null) {
    $allMPR = Export-FIMConfig -OnlyBaseResources -CustomConfig "/ManagementPolicyRule"
    $allWF = Export-FIMConfig -OnlyBaseResources -CustomConfig "/WorkflowDefinition"
    $allSet = Export-FIMConfig -OnlyBaseResources -CustomConfig "/Set"
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

function New-MPRObject()
{
    param ($mpr)

    $set = new-object PSObject

    foreach ($attr in $mpr.ResourceManagementObject.ResourceManagementAttributes) {
        if ($attr.IsMultiValue -eq "True") {
            $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Values
        }
        else
        {
            $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Value
        }
    }

    return $set
}


function New-WFObject()
{
    param ($IdStr)

    $set = new-object PSObject
    $theSet = $allWF | where { $_.ResourceManagementObject.ResourceManagementAttributes.value -eq $IdStr }

    foreach ($attr in $theSet.ResourceManagementObject.ResourceManagementAttributes) {
        if ($attr.IsMultiValue -eq "True") {
            $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Values
        }
        else
        {
            $set | add-member -type NoteProperty -Name $attr.AttributeName -Value $attr.Value
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

#Read All MPRs
function GetAllMPRInformation {

    $items = @()
    foreach ($mpr in $allMPR) {
   
        $theMPR =   New-MPRObject $mpr

        $item = new-object PSObject
        $item | add-member -type NoteProperty -Name "mpr" -Value $theMPR
        $item | add-member -type NoteProperty -Name "wf" -Value @(($theMPR.AuthorizationWorkflowDefinition + $theMPR.AuthenticationWorkflowDefinition + $theMPR.ActionWorkflowDefinition) | % { New-WFObject $_ })

        if ($theMPR.ManagementPolicyRuleType -in ("SetTransition")) {
         
            if ($theMPR.ActionType -eq "TransitionOut") { $theSetID = $theMPR.ResourceCurrentSet } else { $theSetID = $theMPR.ResourceFinalSet }
         
            $item | add-member -type NoteProperty -Name "mprType" -Value $theMPR.ActionType[0]
            $item | add-member -type NoteProperty -Name "TransSet" -Value (New-SetObject $theSetID)
        }

        if ($theMPR.ManagementPolicyRuleType -in ("Request")) 
        {
            
            if ($theMPR.GrantRight -eq $true) 
            {
                $item | add-member -type NoteProperty -Name "mprType" -Value "grantsRight"
            }
            else 
            {
                $item | add-member -type NoteProperty -Name "mprType" -Value "Request"
            }

            if ($theMPR.PrincipalSet -ne $null) 
            {
                $item | add-member -type NoteProperty -Name "ReqSet" -Value (New-SetObject $theMPR.PrincipalSet)
            }
            else
            {
                $item | add-member -type NoteProperty -Name "RelativeToResource" -Value $theMPR.PrincipalRelativeToResource
            }

            $item | add-member -type NoteProperty -Name "CurrentSet" -Value (New-SetObject $theMPR.ResourceCurrentSet)
            $item | add-member -type NoteProperty -Name "FinalSet" -Value (New-SetObject $theMPR.ResourceFinalSet)

        }

        $items += $item
    }

    return $Items

}

function GetActionsNameFromWF {

    PARAM($wf)
    END 
    {
        $actions = @()
        $actionsInXML = ([xml]$wf.XOML).FirstChild.ChildNodes
        $actionsInXML | ForEach-Object {
            if ($_.ActivityDisplayName -ne $null) {
                $actions += $_.ActivityDisplayName
            }
            elseif ($_.Description -ne $null)
            {
                $actions += $_.Description
            }
            elseif ($_.Description -ne $null)
            {
                $actions += $_.Description
            }
            elseif ($_.SynchronizationRuleId -ne $null)
            {
                $actions += "synchronization rule"
            }
            elseif ($_.FunctionExpression -ne $null)
            {
                $actions += "Function Expression Rule"
            }
            elseif ($_.EmailTemplate -ne $null)
            {
                $actions += "Email Template"
            }
            elseif ($_.name -ne $null) 
            {
                $actions += $_.name
            }
            else
            {
                $actions += "(Undefined name)"
            }
        }
        return $actions
    } 
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
<link href="MPR_Description.css" rel="stylesheet">
</head>

<body>

<div class='title'>Management Policy Rules</div>

<div class='title_type'>MPR Types:</div>

<div class="legend grantsRight"><a href="#grantsRight">Grant permissions MPRs</a></div>
<div class="legend SetTransitionIN"><a href="#TransitionIn">Set Transition IN MPRs</a></div>
<div class="legend SetTransitionOUT"><a href="#TransitionOut">Set Transition OUT MPRs</a></div>
<div class="legend Request"><a href="#Request">Request MPRs</a></div>


"@ }
 
            "transitionSet" {   #0=In/Out, 1=Set Name, 2=XPath
                return @"
<tr>
    <td class="SetTransition{0} if left">if object making a transition {0} the set</td>
    <td class="SetTransition{0} if right"><div class="set_title">{1}</div><div class="set_xpath">{2}</div></td>
</tr>
"@ }

            "belong" {  
                return @"
<tr>
    <td class="{0} belong left">and the object belong to this set</td>
    <td class="{0} belong right"><div class="set_title">{1}</div><div class="set_xpath">{2}</div></td>
</tr>
"@ }

            "willBelong" { 
                return @"
<tr>
    <td class="{0} willBelong left">and the object WILL belong to this set</td>
    <td class="{0} willBelong right"><div class="set_title">{1}</div><div class="set_xpath">{2}</div></td>
</tr>
"@ }

            "operation" {  
                return @"
<tr>
    <td class="{0} operation left">doing operation(s)</td>
    <td class="{0} operation right"><div class="operations">{1}</div> <div><span class="operation_sep">{2} </span><span class="operations_attr">{3}</span></div></td>
</tr>
"@ }

            "requestor" {  
                return @"
<tr>
    <td class="{0} requestor left">If someone in this set</td>
    <td class="{0} requestor right"><div class="set_title">{1}</div><div class="set_xpath">{2}</div></td>
</tr>
"@ }

            "relative" { 
                return @"
<tr>
    <td class="{0} relative left">If something related to resource attribute</td>
    <td class="{0} relative right"><div class="resource_attribute">{1}</div></td>
</tr>
"@ }

            "separator" { 
                return @"
<tr>
    <td class="separator" colspan=2></td>
</tr>
"@ }

            "workflow" { 
                return @"
<tr>
    <td class="{0} workflow left">Do {1} workflow</td>
    <td class="{0} workflow right"><div class="wf_title">{2}</div><div class="wf_actions"><ul>{3}</ul></div></td>
</tr>
"@ }

            "grantsRight" { 
                return @"
<tr class="grantsRight grantsRight">
    <td class="grantsRight left"><div>So</div></td>
    <td class="grantsRight left"><div class="txt_grantsRight">Grants Right</div></td>
</tr>
"@ }

            "mprTitle" {   # 0=Class (MPR Type), 1=Status, 2=Name, 3="DISABLED" or $null, 4=Description 
                return @"
<tr>
    <td class="{0} mprTitle"  colspan=2>
        <div class="disabled_{1} mprDN">{2}{3}</div><div class="disabled_{1} mprDescription">{4}</div>
    </td>
</tr>
"@ }

            "sectionTitle" {  
                return @"
<div class='section {0}'><a id='{0}'>{1}</a></div>
"@ }
        }
    }
}

$items = GetAllMPRInformation
$items = $items| Sort-Object -Property mprType
$section = ""

#Header
WriteToFile (HTMLCode "header") "begin"

WriteToFile ""

#region Document all Transition in/out
foreach ($item in $items) {

    "Documenting MPR {0}" -f  $item.mpr.DisplayName

    $disabled = ""
    if ($item.mpr.Disabled -eq "True") { 
        $disabled = " (DISABLED)" 
    } 

    if ($item.mprType -ne $section) {
        if ($section -ne "") {
            WriteToFile "</table>"
        }
        $section = $item.mprType

        switch($section) {
            "grantsRight" { $sectionTXT = "Grant permissions MPRs" } 
            "TransitionIn" { $sectionTXT = "Set Transition IN MPRs" } 
            "TransitionOut" { $sectionTXT = "Set Transition OUT MPRs " } 
            "Request" { $sectionTXT = "Request MPRs" } 
        }

        WriteToFile ((HTMLCode "sectionTitle") -f @($section, $sectionTXT))
        WriteToFile "<table>"

    }

    WriteToFile ((HTMLCode "mprTitle") -f @($item.mprType, $item.mpr.Disabled, $item.mpr.DisplayName, $disabled, $item.mpr.Description))

    if ($item.mpr.ActionType -like "Transition*") { 
        if ($item.mpr.ActionType -eq "TransitionIn") { $InOut = "IN" } else { $InOut = "OUT" }
        if (($item.TransSet.Filter -eq $null) -or ($item.TransSet.Filter -eq "")) { $filter = "(Manually Managed set)" } else { $filter = $item.TransSet.Filter }

        WriteToFile ((HTMLCode "transitionSet") -f @($InOut, $item.TransSet.DisplayName, $filter))
    } 

    if ($item.mprType -in @("Request", "grantsRight")) {

        if ($item.ReqSet -ne $null) {
            WriteToFile ((HTMLCode "requestor") -f @( $item.mprType, $item.ReqSet.DisplayName, $item.ReqSet.Filter ))
        } else {
            WriteToFile ((HTMLCode "relative") -f @( $item.mprType, $item.RelativeToResource, "???" ))
        }

         $operations = $item.mpr.ActionType -join ", "
         if ($item.mpr.ActionParameter -eq "*") {
            $attributes = "all attributes"
            $separator = "on"
         } else
         {
            $attributes = ($item.mpr.ActionParameter -join ", ")
            $separator = "on attribute(s)"
         }
         WriteToFile ((HTMLCode "operation") -f @( $item.mprType, $operations, $separator, $attributes ))
        
         if ($item.CurrentSet.DisplayName -ne $null) {
            WriteToFile ((HTMLCode "belong") -f @( $item.mprType, $item.CurrentSet.DisplayName, $item.CurrentSet.Filter ))
         }

         if ($item.FinalSet.DisplayName -ne $null) {
            WriteToFile ((HTMLCode "willBelong") -f @( $item.mprType, $item.FinalSet.DisplayName, $item.FinalSet.Filter ))
         }

    }

    if ($item.mpr.GrantRight -eq $true) {
        WriteToFile (HTMLCode "grantsRight")
    }
    else
    {
        $actionsStr = ""
        foreach ($wf in $item.wf) {
            $wfDN = $wf.DisplayName
            $wfKind = $wf.RequestPhase
            $actions =  GetActionsNameFromWF $wf
            $actions | ForEach-Object { $actionsStr += "<li>" + $_ + "</li>" }
            WriteToFile ((HTMLCode "workflow") -f @( $item.mprType, $wf.RequestPhase, $wf.DisplayName, $actionsStr))    #0=WF Type, 1=WF Name, 2=actions
        }
    }


    WriteToFile (HTMLCode "separator")
}
#endregion


WriteToFile "</table></body>"
