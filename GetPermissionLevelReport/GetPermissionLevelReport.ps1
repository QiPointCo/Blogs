[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.PowerShell")

function Get-CurrentScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}


#Logs and prints messages
function LogMessage([String] $Msg) {
    Write-Host $Msg -ForegroundColor Green
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") Message: $Msg" | Out-File -FilePath $currentLogPath -Append -Force
}
Function Invoke-LoadMethod() {
    param(
        [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
        [string]$PropertyName
    ) 
    $ctx = $Object.Context
    $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
    $type = $Object.GetType()
    $clientLoad = $load.MakeGenericMethod($type) 


    $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
    $Expression = [System.Linq.Expressions.Expression]::Lambda(
        [System.Linq.Expressions.Expression]::Convert(
            [System.Linq.Expressions.Expression]::PropertyOrField($Parameter, $PropertyName),
            [System.Object]
        ),
        $($Parameter)
    )
    $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
    $ExpressionArray.SetValue($Expression, 0)
    $clientLoad.Invoke($ctx, @($Object, $ExpressionArray))
}
$currentPath = Get-CurrentScriptDirectory
$currentDateTime = Get-Date -format "yyyy-MM-d.hh-mm-ss"

#Path of Log File in Drive.
$currentLogPath = $currentPath + "\" + "PermissionLevel_" + $currentDateTime + ".txt"
LogMessage ("Logging at location: " + $currentLogPath)

#Permission Report
$csvPath = "$currentPath\Site_Permission_Level_Report_" + $currentDateTime + ".csv"
set-content $csvPath "Site Name,Site URL,Group Name,Permission Level"


# Initialize client context
$siteUrl = 'siteurl'
$username = 'username'
$password = 'password'

$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$clientContext.Credentials = $credentials

$Web = $clientContext.Web;
$clientContext.Load($Web)            
$clientContext.Load($Web.RoleAssignments)				
$clientContext.ExecuteQuery()
$rootWebTitle = $Web.Title;
$Url = $Web.Url;
LogMessage ("Looping through : " + $rootWebTitle)
Write-host "Iterating for "$rootWebTitle
$Web.RoleAssignments | % { 
    $clientContext.Load($_.RoleDefinitionBindings)
}
$Web.RoleAssignments | % { 
    $clientContext.Load($_.Member)
}

$clientContext.ExecuteQuery()
$Web.RoleAssignments | % {
    $loginName = $_.Member.LoginName
    $permissionLevel = $_.RoleDefinitionBindings.Name
    add-content $csvPath "$rootWebTitle,$Url,$loginName,$permissionLevel"
}
Write-host "Iterating completed for "$rootWebTitle
LogMessage ("Looping completed for : " + $rootWebTitle)
$Webs = $Web.Webs;
$clientContext.Load($Webs)
$clientContext.ExecuteQuery() 
$level = 1;

function RecursiveWebs($clientContext, $web, $l) {
    $l++;
    $start = "";
    for ($i = 0 ; $i -le $l ; $i++) {
        $start += "  ";
    }

    $counter = 1;
    foreach ($w in $web.Webs) {
        $url = $start + $l + "." + $counter + ": " + $w.Url;

        $clientContext.Load($w.Webs);
        $clientContext.ExecuteQuery();
        if ($w.Webs.Count -gt 0) {

            foreach ($Subweb in $w.Webs) {
			
                Write-host "Iterating for "$Subweb.url				
                $clientContext.Load($Subweb)              

                $clientContext.Load($Subweb.RoleAssignments)				
                $clientContext.ExecuteQuery()
                $subWebTitle = $Subweb.Title;
                $subWebUrl = $Subweb.Url;
                LogMessage ("Looping through : " + $subWebTitle)

                $Subweb.RoleAssignments | % { 
                    $clientContext.Load($_.RoleDefinitionBindings)
                }
                $Subweb.RoleAssignments | % { 
                    $clientContext.Load($_.Member)
                }

                $clientContext.ExecuteQuery()

                $Subweb.RoleAssignments | % {
                    $loginName = $_.Member.LoginName
                    $permissionLevel = $_.RoleDefinitionBindings.Name
                    add-content $csvPath "$subWebTitle,$subWebUrl,$loginName,$permissionLevel"
				
                }				
                Write-host "Iterating completed for "$Subweb.url
                LogMessage ("Looping completed for Subsite : " + $subWebTitle)
            }
            RecursiveWebs $clientContext $w $l
        }
        $counter++
    }
}

if ($web.Webs.Count -gt 0) {
    LogMessage ("Subsites found, loopinf through the subsites....")
    RecursiveWebs $clientContext $Web $level
}