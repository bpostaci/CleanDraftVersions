# THIS CODE-SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.

#Information:
#PowerShell script for deleting older draft versions of Items in a Library.
#Developed by bpostaci, December 2017

#USAGE:
# .\CleanDraftVersions.ps1 -siteUrl <Site Url> -webName <webName> -listName <ListName>  -versionsToKeep <Number> -delete=<$true|$false>
#PARAMETERS
#     -siteUrl       : Url of the Site Collection
#     -webName       : Specific subsite (leave empty for root web)
#     -listName      : Specific ListName
#     -versionsToKeep: Number of major item and its drafts will be preserved (including current version)
#     -delete        : $false (default) for reporting only ,$true for deleting.
#  Highligts: Red will be deleted , Yellow will be preserved , Green current Item .

#EXAMPLE
# .\CleanDraftVersions.ps1 -siteUrl "http://contoso.com" -webName "" -listName "Pages"  -versionsToKeep 10 -delete $false
 
 param (
    [string]$siteUrl = "http://contososp",
    [string]$webName = "",
    [string]$listName = "Pages",
    [int]$versionstoKeep =5,
    [bool]$delete = $false
 )


$site = get-spsite $siteUrl
$web = $site.OpenWeb($webName)
$list = $web.Lists[$listName]


#Add SharePoint PowerShell SnapIn if not already added 
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

$Logfile = ".\Log.txt"


function Write-Log()
{ 
       param
    (
        [Parameter(Position=0,ValueFromPipeline=$true)]
        [string]$msg,
        [string]$ForegroundColor = "White"
    )
    
      $msg | Out-File $Logfile -Append;  
}


Write-Host $siteUrl "Web:" $webName "List:" $listName -ForegroundColor Cyan
Write-Log( $siteUrl + " Web:" + $webName + " List:" + $listName )

$ItemCol = $list.Items

foreach($Item in $ItemCol)
{


$currentVersionsCount= $Item.Versions.count
Write-Host "[Reporting]"
Write-Host "ItemID:" $Item.Id "Version Count:" $currentVersionsCount "Versions To Keep:" $versionstoKeep
Write-Log ("ItemID:" + $Item.Id  + "Version Count:" + $currentVersionsCount + " Versions To Keep:" + $versionstoKeep)


$VersionCol = $Item.Versions;

$arr = New-Object System.Collections.ArrayList


$pubCount =0

    for($i=0; $i -lt $currentVersionsCount ; $i++)
    {
         $itemVer = $VersionCol[$i]

        if( $itemVer.Level -eq "Published")
        {
            $pubCount++
        }

        if($itemVer.Level -eq "Draft" -and $pubCount -ge $versionstoKeep -and $itemVer.IsCurrentVersion -eq 0 )
        {
            Write-Host " -> **ItemID:" $Item.Id "Version Index:" $i "VersionId:"  $itemVer.VersionId " Level:"  $itemVer.Level  "Version Label:" $itemVer.VersionLabel -ForegroundColor Red
            Write-Log (" -> **ItemID:" + $Item.Id  + " Version Index:" + $i + " VersionId:" + $itemVer.VersionId  + " Level:" + $itemVer.Level +  " Version Label:" + $itemVer.VersionLabel)
            if($delete -eq $true)
            {
                $xx = $arr.Add($itemVer.VersionId)
            
            }
        }
        else
        {
            if($ItemVer.IsCurrentVersion -eq 1)
            {
                Write-Host " -> ItemID:" $Item.Id "Version Index:" $i "VersionId:" $itemVer.VersionId "Version Label:" $itemVer.VersionLabel " (*)" -ForegroundColor Green
                Write-Log (" -> ItemID:" + $Item.Id  + " Version Index:" + $i + " VersionId:" + $itemVer.VersionId  + " Level:" + $itemVer.Level +  " Version Label:" + $itemVer.VersionLabel + " (*)")
            }
            else
            {
                Write-Host " -> ItemID:" $Item.Id "Version Index:" $i "VersionId:" $itemVer.VersionId " Level:"  $itemVer.Level "Version Label:" $itemVer.VersionLabel -ForegroundColor Yellow
                Write-Log (" -> ItemID:" + $Item.Id  + " Version Index:" + $i + " VersionId:" + $itemVer.VersionId  + " Level:" + $itemVer.Level +  " Version Label:" + $itemVer.VersionLabel)
            }
            
        }
        
    }


if($delete -eq $true)
{
    Write-Host "[Deleting]"
    foreach($x in $arr)
    {
        try
        {
            $v = $Item.Versions.GetVersionFromID($x)
            Write-Host " -> DELETED ! ItemID:" $Item.Id "Version Index:" $x "VersionId:"  $v.VersionId " Level:"  $v.Level  "Version Label:" $v.VersionLabel 
            Write-Log(" -> DELETED ! ItemID:" + $Item.Id + "Version Index:" + $x + "VersionId:" + $v.VersionId + " Level:" +  $v.Level + "Version Label:" + $v.VersionLabel)
            $v.Delete()
        }
        catch
        {
            Write-Host "An error occured ! index=" $x "Message" $_.Message -ForegroundColor Red
            Write-Log("An error occured ! index=" + $x + "Message" + $_.Message) 
        }
    }
    
}
 
}