[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.PowerShell")


#Gets the current directory.
function Get-CurrentScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}
function LogWarning([String] $ErrorMsg)
{
    Write-Warning $ErrorMsg
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") WARNING: $ErrorMsg" | Out-File -FilePath $currentLogPath -Append -Force
}

#Logs and prints error messages
function LogError([String] $ErrorMessage, [String]$ErrorDetails, [String]$ErrorPosition)
{
    Write-Host $ErrorMessage -foregroundcolor red
    $fullErrorMessage = $ErrorMessage + $ErrorDetails + ". " + $ErrorPosition
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ERROR: $fullErrorMessage" | Out-File -FilePath $currentLogPath -Append -Force 
}

$currentPath = Get-CurrentScriptDirectory
$currentDateTime = Get-Date -format "yyyy-MM-d.hh-mm-ss"

#Path of Log File in Drive.
$currentLogPath = $currentPath + "\" + "CreateSPOnlineLogs_"+ $currentDateTime +".txt"
function CreateLists
{
    Param ($siteNode, $baseUrl)

    
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($baseUrl)
    $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword)      
    $Context.Credentials = $Creds

    
     foreach ($list in $siteNode.List) 
    {
        Write-Host "Creating List "  $list.Name " in the subsite " $subsiteUrl -foregroundcolor "magenta"
        try{
        $web = $Context.Web
        $templates = $context.web.listtemplates
        $Context.Load($templates)
        $Context.ExecuteQuery()

        $template = $templates | Where-Object{ $_.FeatureId -eq $list.TemplateFeatureId }

        $lci = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $lci.Title = $list.Name
        $lci.TemplateFeatureId = $template.FeatureId
        $lci.TemplateType = $template.ListTemplateTypeKind
        $lci.DocumentTemplateType = $template.ListTemplateTypeKind

        $lists = $Context.Web.Lists;
        $Context.Load($lists);
        $Context.ExecuteQuery();

        $list = $lists.Add($lci)
        $list.Update()
        $Context.ExecuteQuery()

        }catch{
            if($_.Exception.Message -like '*A list, survey, discussion board, or document library with the specified title already exists in this Web site*' )
		    {
                LogWarning "List already exists .. "
			        
		    }else{
                LogError "Error creating lists " $_.Exception.Message $_.Exception.GetType().FullName $_.InvocationInfo.PositionMessage
            }
        }
    }
}
function Enable-Feature
{
	Param ($siteNode, $baseUrl)
    
    
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($baseUrl)
    $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword)      
    $Context.Credentials = $Creds
            
    
    foreach ($feature in $siteNode.Feature) 
    {
        Write-Host "Activating Feature " $feature.Id " for site " $subsiteUrl -foregroundcolor black -backgroundcolor yellow
        try{
            $featureGuid = new-object System.Guid $feature.Id
		
	        $features = $null	
	
	        if ($feature.Scope -eq [Microsoft.SharePoint.Client.FeatureDefinitionScope]::Site)
	        {
	
		        $features = $Context.Site.Features
		
	        } else {
	
		        $features = $Context.Web.Features
		
	        }
	        $Context.Load($features)
	        $Context.ExecuteQuery()
	
	        $createdfeature = $features.Add($featureGuid, $force, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
	
	        # TODO: Check if the feature is already enabled
	        $Context.ExecuteQuery()
        
        }catch{
             
             if($_.Exception.Message -like '*is already activated at scope*' -or $_.Exception.Message -like '*(407) Proxy Authentication Required*')
		        {
                    LogWarning "Feature is already activated .. "
			        
		        }else{
                    LogError "Error Activating Feature " $_.Exception.Message $_.Exception.GetType().FullName $_.InvocationInfo.PositionMessage
                }
        }
        
	
	    
    }	
	
} 
function CreateSites($baseUrl, $sites, [int]$progressid) 
{      


    $sitecount = $sites.ChildNodes.Count 
    $counter = 0 

    foreach ($site in $sites.Site) 
    { 
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($baseUrl)
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword)
        $Context.Credentials = $Creds

        Write-Progress -ID $progressid -Activity "Creating sites" -status "Creating $($site.Name)" -percentComplete ($counter / $sitecount*100) 
        $counter = $counter + 1

        Write-Host "Creating $($site.Name) $($baseUrl)/$($site.Url)" -foregroundcolor Blue -backgroundcolor white
        
        ##New-SPWeb -Url "$($baseUrl)/$($site.Url)" -AddToQuickLaunch:$false -AddToTopNav:$false -Confirm:$false -Name "$($site.Name)" -Template $site.Template -UseParentTopNav:$true 
        try{

            $WCI = New-Object Microsoft.SharePoint.Client.WebCreationInformation
            $WCI.WebTemplate = $site.Template
            $WCI.Title = $site.Name
            $WCI.Url = $site.Url
            
            $WCI.Language = "1033"
           
            $subWeb = $Context.Web.Webs.Add($WCI)       
            $subWeb.BreakRoleInheritance($false, $false);         
            $subWeb.Update() 
            $Context.Load($subWeb)
            $Context.ExecuteQuery()
        
        }catch{
                
            if($_.Exception.Message -like '*is already in use.*' -or $_.Exception.Message -like '*(407) Proxy Authentication Required*')
		    {
                LogWarning "Site already exists .. "
			        
		    }else{
                LogError "Error Creating Sites " $_.Exception.Message $_.Exception.GetType().FullName $_.InvocationInfo.PositionMessage
            }
        }
        
        $subsiteUrl = $Context.Url + "/" + $site.Url

        Enable-Feature -siteNode $site -baseUrl $subsiteUrl

        CreateLists -siteNode $site -baseUrl $subsiteUrl

        if ($site.ChildNodes.Count -gt 0) 
        { 
            CreateSites "$($baseUrl)/$($site.Url)" $site ($progressid +1) 
        }
        
         
        Write-Progress -ID $progressid -Activity "Creating sites" -status "Creating $($site.Name)" -Completed 
    } 
    
    

} 

# read an xml file 
$xml = [xml](Get-Content "demosites.xml") 
$xml.PreserveWhitespace = $false

# Initialize client context
$siteCollectionUrl = '<Site Collection Url>'
$username = '< Admin Username>'
$password = '< Admin Password>' 
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force    
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($baseUrl)
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username,$securePassword)
$Context.Credentials = $Creds
    

Enable-Feature -siteNode $xml.Sites -baseUrl $siteCollectionUrl
CreateLists -siteNode $xml.Sites -baseUrl $siteCollectionUrl
CreateSites $siteCollectionUrl $xml.Sites 1 