$totalSite = 1

function CreateSubSites-Helper
{
    param(
        [string] $sUrl,
        [string] $levelUrl,
        [string] $lineUrl,
        [string] $lineLevelSite,
        [string] $lineLevelSiteTemplate,
        [string] $linePermInheritance,
        [string] $lineCopyPermsFromParent,
        [string] $webUrl,
        [string] $siteUrl,
        [Object] $line,
        [string] $lineInheritNavigation,
        [bool] $IsRootWeb,
        [int] $totalNoSite
    )

    $localcontext = New-Object Microsoft.SharePoint.Client.ClientContext($sUrl) 
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
    $localcontext.Credentials = $credentials 
    $newWeb = $localcontext.Web
    $localcontext.Load($newWeb)
    try
    {
    $localcontext.ExecuteQuery()
    $siteExists =$true
    $levelUrl = $newWeb.Url
    $localcontext.Dispose()
    }
    catch
    {
        $siteExists =$false
        $localcontext.Dispose()
    }

    if($siteExists)
    {
        Write-Host "=====================================================================================" -ForegroundColor Yellow

        Write-Host "Site Already Created - $sUrl - Total Number of Sites Created So far - " $totalSite  "Current Level - " $levelUrl -ForegroundColor Green

        Write-Host "=====================================================================================" -ForegroundColor Yellow

        $siteExists = $false
        $context.Dispose()
    }
    else
    {
        $displayDetails = $levelUrl+"/"+ $lineUrl
        Write-Host "Creating subsite -" $lineLevelSite " at: " $displayDetails  -foregroundcolor yellow


        $context = New-Object Microsoft.SharePoint.Client.ClientContext($levelUrl) 
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
        $context.Credentials = $credentials 
        $web = $context.Web    
        $site = $context.Site 

        $context.Load($web)
        

        $wci = New-Object Microsoft.SharePoint.Client.WebCreationInformation
        $wci.Url = $lineUrl
        $wci.Title = $lineLevelSite
        $wci.WebTemplate = $lineLevelSiteTemplate

        if($IsRootWeb)
        {

          $newWeb = $site.RootWeb.Webs.Add($wci);
        }
        else
        {
            $newWeb = $context.Web.Webs.Add($wci); 
        }

        $context.Load($newWeb)

        try
        {
            $context.ExecuteQuery();
        }
        catch
        {
            Write-Host "$siteUrl $_.Exception.Message" -foregroundcolor red
        }

        if($linePermInheritance -eq $false)
        {
            if ($lineCopyPermsFromParent -eq $false)
            {
                $newWeb.BreakRoleInheritance($false, $false)
            }
            else
            {
                $newWeb.BreakRoleInheritance($true, $true)
                write-host "Breaking inheritance from parent " $webUrl -ForegroundColor Green
            }


            $columnNumber = 1
            $columnAccountType = "AccountType" + $columnNumber
            $columnAccountName = "AccountName" + $columnNumber
            $columnPermLevel = "PermLevel" + $columnNumber
            $addUsers = "AddUsers"+ $columnNumber

            while ($line.$columnAccountType)
            {
                $account = $null
                #Check to see what type of account object is to be assigned - e.g., SharePoint group, AD account, etc.
                if ($line.$columnAccountType -eq "SPGroup")
                {
                    #Check to see if SharePoint group exists in the site collection
                    $context.Load($web.SiteGroups)
                    $context.ExecuteQuery()

                    $gName = $web.SiteGroups | where {$_.title -eq $line.$ColumnAccountName}

                    if (-not $gName)
                    {
                        #Create SharePoint group
                        $account = New-Object Microsoft.SharePoint.Client.GroupCreationInformation
                        $account.Title = $line.$ColumnAccountName
                        $account.Description = "Custom SharePoint Group"
                        $account = $newWeb.SiteGroups.Add($account)
                        $context.Load($account);
                        $newWeb.Update()
                        $context.ExecuteQuery()

                        $grpTitle = $account.Title

                        Write-Host "Created group $grpTitle succesfully!" -foregroundcolor yellow 

                        #Add Users to the groups
                        $usersCSV = $line.$AddUsers
                        [string[]]$usersObjCount = $usersCSV.Split(",")
                        $usersCount = $usersObjCount.count

                        $usersObjCount = $usersObjCount | select -uniq

                        if($usersCount -ge 1)
                        {
                            foreach ($user in $usersObjCount)
                            {
                                if(-not [string]::IsNullOrEmpty($user)) 
                                {  
                                    $spoUser = $context.Web.EnsureUser($user) 
                                    $context.Load($spoUser) 
                                    $spoUserToAdd=$account.Users.AddUser($spoUser) 
                                    $context.Load($spoUserToAdd) 
                                    $context.ExecuteQuery()    
                                    Write-Host "SharePoint User $user added succesfully!" -foregroundcolor Green 
                                }
                            }
                        }
                    }
                    else 
                    {
                        $account = $web.SiteGroups.GetByName($line.$ColumnAccountName)
                        $context.Load($account);
                        $context.ExecuteQuery()
                    }
                }
                else
                {
                    #Set account variable to SPUser
                    $account = $newWeb.EnsureUser($line.$ColumnAccountName)
                    $context.Load($account);
                    $context.ExecuteQuery()
                }

                $access = $newWeb.RoleDefinitions.GetByName($line.$ColumnPermLevel)  
                $roleAssignment =  New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)  

                $roleAssignment.Add($access)  
                $context.Load($newWeb.RoleAssignments.Add($account, $roleAssignment))  
                $newWeb.Update()
                $context.ExecuteQuery()

                #Set up the next account column number to view from CSV file
                $columnNumber = $columnNumber + 1
                $columnAccountType = "AccountType" + $columnNumber
                $columnAccountName = "AccountName" + $columnNumber
                $columnPermLevel = "PermLevel" + $columnNumber
                $addUsers = "AddUsers"+ $columnNumber
            }
        }
        else
        {
            $newWeb.ResetRoleInheritance()
            write-host "Inheriting Permission from parent " $newWeb.Url -ForegroundColor Green
        }

        if($line.InheritNavigation -eq $true)
        {
            $newWeb.Navigation.UseShared = $true;
            Write-Host "Navigation is inherting from parent"  $newWeb.Url -ForegroundColor Green
        }

        $levelUrl = $newWeb.Url

        $context.Load($newWeb)
        $context.ExecuteQuery();
        $context.Dispose();
        Write-Host "Site created:" $newWeb.Title " at: " $displayDetails -foregroundcolor green 

        Write-Host "=====================================================================================" -ForegroundColor Yellow

        Write-Host "Total Number of Sites Created So far - " $totalSite -ForegroundColor Green
                 
        Write-Host "=====================================================================================" -ForegroundColor Yellow
    }
}



function New-SPSitesFromCsv
{
    Param (
    [parameter(Mandatory=$true)][string]$CsvFile        
    )
     
    try
    {
        $level1Site = $null
        $level2Site = $null
        $level3Site = $null
        $level4Site = $null
        $level5Site = $null
        $siteExists = $false
        

        $url ="Mention your Office 365 site collection url"
        $username="Mention your user name@onmicrosoft.com"

        $password = read-host -prompt "Password for $username" -AsSecureString

        Set-Location $PSScriptRoot
        Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
        Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 
        
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
        $Context.Credentials = $credentials 
        $spWweb = $context.Web
        $spSite = $context.Site 
        
        $context.Load($spWweb)
        $context.Load($spSite)
        try
        {
            $context.ExecuteQuery()
        }
        catch
        {
            Write-Host "$url $_.Exception.Message" -foregroundcolor red
            return
        }

        $csvData = Import-Csv $CsvFile
        foreach ($line in $csvData)
        {
            if ($Level1Site -eq $line.Level1Site)
            {               
                if ($level2Site -eq $line.Level2Site)
                {
                    if ($Level3Site -eq $line.Level3Site)
                    {
                        if ($Level4Site -eq $line.Level4Site)
                        {
                            # level 5 check
                            $sUrl = $Level4Url + "/" + $line.url
                            CreateSubSites-Helper $sUrl $Level4Url $line.Url $line.Level5Site $line.SiteTemplate $line.PermInheritance $line.CopyPermsFromParent $spWweb.Url $url $line $line.InheritNavigation $false $totalSite
                            $Level5Site = $line.Level5Site
                            $Level5Url = $sUrl
                            $totalSite = $totalSite +1
                        }
                        else
                        {
                            # level 4 check
                            $sUrl = $Level3Url + "/" + $line.url
                            CreateSubSites-Helper $sUrl $Level3Url $line.Url $line.Level4Site $line.SiteTemplate $line.PermInheritance $line.CopyPermsFromParent $spWweb.Url $url $line $line.InheritNavigation $false $totalSite
                            $Level4Site = $line.Level4Site
                            $Level4Url = $sUrl
                            $totalSite = $totalSite +1
                        }  
                    }
                    else
                    {
                        #level 3 check
                        $sUrl = $Level2Url + "/" + $line.url
                        CreateSubSites-Helper $sUrl $Level2Url $line.Url $line.Level3Site $line.SiteTemplate $line.PermInheritance $line.CopyPermsFromParent $spWweb.Url $url $line $line.InheritNavigation $false $totalSite
                        $Level3Site = $line.Level3Site
                        $Level3Url = $sUrl
                        $totalSite = $totalSite +1
                    }  
                }
                else
                {
                    #Level 2 Check
                    $sUrl = $Level1Url +"/" + $line.url
                    CreateSubSites-Helper $sUrl $Level1Url $line.Url $line.Level2Site $line.SiteTemplate $line.PermInheritance $line.CopyPermsFromParent $spWweb.Url $url $line $line.InheritNavigation $false $totalSite
                    $Level2Site = $line.Level2Site
                    $Level2Url = $sUrl
                    $totalSite = $totalSite +1
                }                
            }
            else
            {
                #Level 1 Check
                $sUrl = $spSite.Url + "/" + $line.url
                CreateSubSites-Helper $sUrl $spSite.Url $line.Url $line.Level1Site $line.SiteTemplate $line.PermInheritance $line.CopyPermsFromParent $spWweb.Url $url $line $line.InheritNavigation $true $totalSite
                $Level1Site = $line.Level1Site
                $Level1Url = $sUrl
                $totalSite = $totalSite +1
            }
        }
    }
    catch 
    {
        write-host $_ -foregroundcolor red 
    }
    finally
    {
    }
}

New-SPSitesFromCsv -CsvFile "C:\SitesDemo-Demo.csv"

