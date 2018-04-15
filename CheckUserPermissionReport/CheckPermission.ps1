[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.PowerShell")

function Get-CurrentScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}


#Logs and prints messages
function LogMessage([String] $Msg)
{
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
                [System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),
                [System.Object]
            ),
            $($Parameter)
   )
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}
$currentPath = Get-CurrentScriptDirectory
$currentDateTime = Get-Date -format "yyyy-MM-d.hh-mm-ss"

#Path of Log File in Drive.
$currentLogPath = $currentPath + "\" + "CheckUserPermissionOnSite_"+ $currentDateTime +".txt"

#Permission Report
$csvPath = "$currentPath\Site_Permission_Report_" + $currentDateTime +".csv"
set-content $csvPath "Site Name,List Name, Item Title, User Name, Permission"


# Initialize client context
$siteUrl = 'Site url'
$username = 'admin username'
$password = 'admin password'


$securePassword = ConvertTo-SecureString $password -AsPlainText -Force

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword)
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$clientContext.Credentials = $credentials

$Web = $clientContext.Web;
$clientContext.Load($Web)
$clientContext.ExecuteQuery()


$Url = $Web.Url;
Write-Host $Url;

$SearchUser = Read-Host "Enter user to check permission"

$SearchUser = "i:0#.f|membership|"+$SearchUser 
Write-Host "Searching permission for the user " $SearchUser

$Webs = $Web.Webs;
$clientContext.Load($Webs)
$clientContext.ExecuteQuery() 

$Lists = $Web.Lists
$clientContext.Load($Lists)
$clientContext.ExecuteQuery()
 
#Iterate through each list in a site   
ForEach($List in $Lists)
{
    #Get the List Name
    #Write-host $List.Title
   

    if($List.BaseType -eq "GenericList")
    {
        LogMessage(" InheritedPermissionList ")
        Invoke-LoadMethod -Object $List -PropertyName "HasUniqueRoleAssignments"
        $clientContext.ExecuteQuery()

        # Write-Host $List.HasUniqueRoleAssignments
        

        if($List.HasUniqueRoleAssignments -eq $true)
        {

            $ListTitle = $List.Title
            Write-host $List.Title "has broken permission"

            $RoleAssignments = $List.RoleAssignments;
            $clientContext.Load($RoleAssignments)
            $clientContext.ExecuteQuery()

            foreach($ListRoleAssignment in $RoleAssignments)
            {

                $member = $ListRoleAssignment.Member
                $roleDef = $ListRoleAssignment.RoleDefinitionBindings

                $clientContext.Load($member)
                $clientContext.Load($roleDef)
                $clientContext.ExecuteQuery()

                     #Is it a User Account?
                     if($ListRoleAssignment.Member.PrincipalType -eq "User")   
                     {
                         #Is the current user is the user we search for?
                         Write-Host "Current user : " $ListRoleAssignment.Member.LoginName 
                         if($ListRoleAssignment.Member.LoginName -eq $SearchUser)
                         {
                         #Write-Host  $SearchUser has direct permissions to List ($List.ParentWeb.Url)/($List.RootFolder.Url)
                         #Get the Permissions assigned to user

                          $UserDisplayName = $ListRoleAssignment.Member.Title;
                          $ListUserPermissions=@()
                            foreach ($RoleDefinition  in $roleDef)
                            {
                                 $ListUserPermissions += $RoleDefinition.Name +";"
                            }
                            add-content $csvPath "$Url,$ListTitle,,$UserDisplayName,$ListUserPermissions"
                             #Send the Data to Log file
                             #"$($List.ParentWeb.Url)/$($List.RootFolder.Url) `t List `t $($List.Title)`t Direct Permissions `t $($ListUserPermissions)" | Out-File UserAccessReport.csv -Append
                          }
                    }
            }


            $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
            $camlQuery.ViewXml ="<View Scope='RecursiveAll' />";
            $ListItems= $List.GetItems($camlQuery)
            $clientContext.Load($ListItems)
            $clientContext.ExecuteQuery()

            foreach($item in $ListItems)
            {
                Write-Host "##############"
               
                $itemTitle = $item["Title"]
                Write-Host "Item:" $item["Title"]
                Write-Host "##############"

                Invoke-LoadMethod -Object $item -PropertyName "HasUniqueRoleAssignments"
                $clientContext.ExecuteQuery()
                if ($item.HasUniqueRoleAssignments -eq $true)
                {
                    $itemRoleAssignments = $item.RoleAssignments;
                    $clientContext.Load($itemRoleAssignments)
                    $clientContext.ExecuteQuery()

                    foreach($itemRoleAssignment in $itemRoleAssignments)
                    {

                        $Itemmember = $itemRoleAssignment.Member
                        $ItemroleDef = $itemRoleAssignment.RoleDefinitionBindings

                        $clientContext.Load($Itemmember)
                        $clientContext.Load($ItemroleDef)
                        $clientContext.ExecuteQuery()

                             #Is it a User Account?
                             if($itemRoleAssignment.Member.PrincipalType -eq "User")   
                             {
                                 #Is the current user is the user we search for?
                                 Write-Host "Current Item user : " $itemRoleAssignment.Member.LoginName 
                                 if($itemRoleAssignment.Member.LoginName -eq $SearchUser)
                                 {
                                 Write-Host  $SearchUser has direct permissions to List ($List.ParentWeb.Url)/($List.RootFolder.Url)
                                 #Get the Permissions assigned to user

                                  $UserDisplayName = $itemRoleAssignment.Member.Title;
                                  $ItemUserPermissions=@()
                                    foreach ($ItemRoleDefinition  in $ItemroleDef)
                                    {
                                         $ItemUserPermissions += $ItemRoleDefinition.Name +";"
                                    }
                                    add-content $csvPath "$Url,$ListTitle,$itemTitle,$UserDisplayName,$ItemUserPermissions"
                                     #Send the Data to Log file
                                     #"$($List.ParentWeb.Url)/$($List.RootFolder.Url) `t List `t $($List.Title)`t Direct Permissions `t $($ListUserPermissions)" | Out-File UserAccessReport.csv -Append
                                  }
                            }
                    }
                }

                #Write-Host $item.HasUniqueRoleAssignments
            }

            #Write-Host $ListItems.Count

        }
        else
        {

            Write-host $List.Title "has inherited permission"

        }
    }
}