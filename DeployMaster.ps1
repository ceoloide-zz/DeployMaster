[CmdletBinding()]
Param(
   [switch]$NoGUI,
   [string]$LogFilePath
)

$Version = "0.3 (Beta)"
$Build = 1

#
# Adds PowerShell SharePoint snap-in if not loaded
#
if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
  Add-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue
}

#
# This cmdlet takes care of the disposable objects to prevent memory leak. 
#
Start-SPAssignment -Global

#
# Global variables
#
$Force = $false
$ConfigPath
$WebAppUrl
$WebApp
$ConfName
[xml]$ConfigFile

#
# Displays the splash screen
#
function GUI-SplashScreen()
{
	Clear-Host
	Write-Host " "
	Write-Host "  ____             _               __  __           _             "
	Write-Host " |  _ \  ___ _ __ | | ___  _   _  |  \/  | __ _ ___| |_ ___ _ __  "
	Write-Host " | | | |/ _ \ '_ \| |/ _ \| | | | | |\/| |/ _` / __| __/ _ \ '__| "
	Write-Host " | |_| |  __/ |_) | | (_) | |_| | | |  | | (_| \__ \ ||  __/ |    "
	Write-Host " |____/ \___| .__/|_|\___/ \__, | |_|  |_|\__,_|___/\__\___|_|    "
	Write-Host "            |_|            |___/                                  "
	Write-Host " "
	Write-Host " "
	Write-Host " Version: " -NoNewLine
	Write-Host $Version -ForegroundColor Green -NoNewLine
	Write-Host (" - Build " + $Build)
	Write-Host

	Write-Host " Written by: " -NoNewLine
	Write-Host "Massarelli, Marco" -ForegroundColor Green -NoNewline
	Write-Host " <marco.massarelli@gmail.com>" -ForegroundColor Yellow -NoNewLine
	Write-Host " (Google, Italy)"

	Write-Host " Core functionality: " -NoNewLine
	Write-Host "Calloni, Paolo" -ForegroundColor Green -NoNewline
	Write-Host " <paolo.calloni@avanade.com>" -ForegroundColor Yellow -NoNewLine
	Write-Host " (Avanade, Italy)"
	Write-Host
	Write-Host
}


#
# Waits for the user to press [ENTER]
#
function PressEnterToContinue()
{
	Write-Host ""
	Write-Host "Press [ENTER] to continue"
	Read-Host
}

function Is-Int ($Value) {
    return $Value -match "^[\d]+$"
}

function Is-Numeric ($Value) {
    return $Value -match "^[\d\.]+$"
}

#
# Formats a string with the specified length, fills the rest with the specified characters
#
function FormatString($StringValue, $MaxLength, $Left, $FillChar)
{
	if ($StringValue -eq $null)
	{ $StringValue = "" }
	
	if ($StringValue.Length > $MaxLength)
	{ $StringValue = $StringValue.Substring(0, $MaxLength)}
	
	if ($Left)
	{return $StringValue.PadLeft($MaxLength, $FillChar)}
	else
	{return $StringValue.PadRight($MaxLength, $FillChar)}
}

#
# Prints a choice menu and waits for user inputs, returning the user choice
#
function PrintChoiceMenu($Title, $Choices)
{
	Clear-Host
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	Write-Host $Title
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	for ($i = 0; $i -lt $Choices.Length; $i++)
	{
		Write-Host $Choices[$i]
	}
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	$OpChoice = Read-Host 
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	return $OpChoice
}

#
# Writes a line in the Host preceded by a red ERROR string
#
function Write-Host-Error($message)
{
		$errorStr = FormatString -StringValue "ERROR" -MaxLength 10 -Left $false -FillChar ' '
		Write-Host $errorStr -ForegroundColor Red -NoNewLine
		Write-Host $message
}

#
# Writes a line in the Host preceded by a yellow WARNING string
#
function Write-Host-Warning($message)
{
		$warningStr = FormatString -StringValue "WARNING" -MaxLength 10 -Left $false -FillChar ' '
		Write-Host $warningStr -ForegroundColor Yellow -NoNewLine
		Write-Host $message
}

#
# Writes a line in the Host preceded by a green OK string
#
function Write-Host-OK($message)
{
		$okStr = FormatString -StringValue "OK" -MaxLength 10 -Left $false -FillChar ' '
		Write-Host $okStr -ForegroundColor Green -NoNewLine
		Write-Host $message
}

#
# Copies files from source to destination
#
function CollectFiles($SolutionNodes, $srcPath, $destFolder, $ext, $reldbg)
{
	foreach( $solution in $SolutionNodes) 
	{
		$SolutionId = $solution.GetAttribute("id")
		$SolWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($SolutionId)

		if ( -not (Test-Path $destFolder) ) {New-Item $destFolder  -Type Directory  | Out-Null}
		if ( -not (Test-Path ($destFolder + $reldbg )) ) {New-Item ($destFolder + $reldbg) -Type Directory  | Out-Null}

		$destFile = $destFolder + $reldbg + "\" + $SolWithoutExt + "." + $ext
		$srcFile = $srcPath + "\" + $SolWithoutExt + "\bin\" + $reldbg + "\" + $SolWithoutExt + "." + $ext
		
		Write-Host "Copying $ext file for solution: $SolutionId"
		Copy-Item -LiteralPath $srcFile -Destination  $destFile -Force
	}
}

#
# Waits for the specified solutions to terminate the deploy process.
# Checks for Deploy ($true) or Retract ($false) depending on the $deploying parameter.
# Sleeps for $waitInterval seconds after each check.
#
function WaitDeployProcessForSeconds($solutions, $deploying, $waitInterval)
{
	$errorOccurred = $false
	Write-Host "Waiting for operation to end [" -NoNewline
	
	#
	# Cycle while there is at least one deploy/retract pending
	#
	do
	{
		$Installed = $true
		foreach( $solution in $solutions ) 
		{
			$SPsolution = Get-SPSolution -Identity $solution.GetAttribute("id") -ErrorAction SilentlyContinue 

			if ($SPsolution -eq $null)
			{ continue }
			
			if ($SPsolution.JobExists) 
			{ $Installed = $false }
		} 
		if ($Installed) { break }
		sleep -Seconds $waitInterval
		Write-Host "." -NoNewline
	}
	while($Installed -eq $false)
	
	Write-Host "] Done!"
	Write-Host "Now checking for errors."
	Write-Host

	# 
	# Check the result of the deploy / retract
	#
	$errorOccurred = $false
	foreach( $solution in $solutions) 
	{
		$SPsolution = Get-SPSolution -Identity $solution.GetAttribute("id") -ErrorAction SilentlyContinue
		$solutionName = FormatString -StringValue $solution.GetAttribute("id") -MaxLength 40 -Left $false -FillChar ' '
		Write-Host $solutionName -NoNewline

		if ($SPsolution -eq $null)
		{ 
			Write-Host-Warning -message "Not Existing"
			continue
		}
		
		if (($deploying -and !$SPsolution.Deployed) -or (!$deploying -and $SPsolution.GlobalDeployed))
		{ 
			Write-Host-Error -message $SPsolution.DeploymentState
			$errorOccurred = $true 
		}
		else
		{
			Write-Host-OK -message $SPsolution.DeploymentState
		}
	} 
	Write-Host ""
	if ($errorOccurred)
		{Write-Host "An error occurred during the solution deployment." -ForegroundColor Red}
	else
		{Write-Host "Finished"}
}

#
# Waits for the specified solutions to terminate the deploy process.
# Checks for Deploy ($true) or Retract ($false) depending on the $deploying parameter.
# Sleeps for 30 seconds after each check.
#
function WaitDeployProcess($solutions, $deploying)
{
	WaitDeployProcessForSeconds -solutions $solutions -deploying $deploying -waitInterval 30
}

#
# Waits for the specified solution to terminate the deploy process.
# Checks for Deploy ($true) or Retract ($false) depending on the $deploying parameter.
# Sleeps for $waitInterval seconds after each check.
#
function WaitDeployProcessForSeconds-Single($solutionID, $deploying, $waitInterval)
{
	$errorOccurred = $false
	Write-Host "Waiting for operation to end [" -NoNewline
	
	#
	# Cycle while the deploy/retract is pending
	#
	do
	{
		$Installed = $true

		$SPsolution = Get-SPSolution -Identity $solutionID -ErrorAction SilentlyContinue 

		if ($SPsolution -eq $null)
		{ continue }
			
		if ($SPsolution.JobExists) 
		{ $Installed = $false }

		if ($Installed) { break }
		sleep -Seconds $waitInterval
		Write-Host "." -NoNewline
	}
	while($Installed -eq $false)
	
	Write-Host "] Done!"
	Write-Host "Now checking for errors."
	Write-Host

	# 
	# Check the result of the deploy / retract
	#
	$errorOccurred = $false

	$SPsolution = Get-SPSolution -Identity $solutionID -ErrorAction SilentlyContinue
	$solutionName = FormatString -StringValue $solutionID -MaxLength 40 -Left $false -FillChar ' '
	Write-Host $solutionName -NoNewline

	if ($SPsolution -eq $null)
	{ 
		Write-Host-Warning -message "Not Existing"
	}
	elseif (($deploying -and !$SPsolution.Deployed) -or (!$deploying -and $SPsolution.GlobalDeployed))
	{ 
		Write-Host-Error -message $SPsolution.DeploymentState
		$errorOccurred = $true 
	}
	else
	{
		Write-Host-OK -message $SPsolution.DeploymentState
	}
	 
	Write-Host ""
	if ($errorOccurred)
		{Write-Host "An error occurred during the solution deployment." -ForegroundColor Red}
	else
		{Write-Host "Finished"}
}

#
# Waits for the specified solution to terminate the deploy process.
# Checks for Deploy ($true) or Retract ($false) depending on the $deploying parameter.
# Sleeps for 30 seconds after each check.
#
function WaitDeployProcess-Single($solutionID, $deploying)
{
	WaitDeployProcessForSeconds-Single -solutionID $solutionID -deploying $deploying -waitInterval 30
}

#
# Restarts Sharepoint TimerJob Service on the farm
#
function OP-RestartOWSTimer() 
{
    Write-Host "Restarting SharePoint TimerJob Service on the farm"
    $farm = get-SPFarm
    $instances = $farm.TimerService.Instances
    foreach ($instance in $instances)
    {
        write-host ("    " + $instance.Server.Address + " ") -NoNewLine
        $instance.Stop()
        Write-Host "Stopped " -ForegroundColor Red -NoNewLine
        $instance.Start()
        Write-Host "Started" -ForegroundColor Green
    }
}

#
# Adds a solution to the solution store in the farm
#
function AddSolution($SolutionId, $SolutionPath) 
{
	$SPsolution = Get-SPSolution -Identity $SolutionId -ErrorAction SilentlyContinue
	if ($SPsolution -eq $null)
	{
		Add-SPSolution $SolutionPath -ErrorAction Continue
	}
	else
	{
		Write-Host-Warning -message ("Solution " + $SolutionId + " was already added.")
	}
}

#
# Adds multiple solutions to the solution store in the farm
#
function AddSolution-Multi($SolutionNodes, $RootPath) 
{
	foreach( $solution in $SolutionNodes) 
	{
		AddSolution -SolutionId $solution.GetAttribute("id") -SolutionPath ($RootPath + $solution.GetAttribute("id"))
	}
	Write-Host
}

#
# Installs a solution from the solution store in the farm
#
function InstallSolution($SolutionId, $GacDeploy, $WebDeploy, $WebApp, $Force)
{
	Write-Host ("Installing solution " + $SolutionId)
	$SPsolution = Get-SPSolution -Identity $SolutionId -ErrorAction SilentlyContinue 
	if ($SPsolution -eq $null)
	{
		Write-Host-Error -message ("Solution " + $SolutionId + " does not exist")
	}
	elseif ($SPsolution.Deployed -or $SPsolution.GlobalDeployed)
	{
		Write-Host-Warning -message ("Solution " + $SolutionId + " is already installed")
	}
	elseif ($WebDeploy -eq $true)
	{
		Install-SPSolution -Identity $SolutionId -GACDeployment:$GacDeploy -CASPolicies:$($SPsolution.ContainsCasPolicy) -WebApplication $WebApp -Confirm:$false -Force:$Force
	}
	else
	{
		Install-SPSolution -Identity $SolutionId -GACDeployment:$GacDeploy -CASPolicies:$($SPsolution.ContainsCasPolicy) -Confirm:$false -Force:$Force
    } 
}

#
# Installs multiple solutions from the solution store in the farm and wait the end of the process
#
function InstallSolution-Multi($SolutionNodes, $RootPath, $WebApp, $Force)
{
	foreach( $solution in $SolutionNodes) 
	{
		[bool] $GacDeploy = [System.Convert]::ToBoolean($solution.GetAttribute("GACdeploy"))
		[bool] $WebDeploy = [System.Convert]::ToBoolean($solution.GetAttribute("webapp"))

		InstallSolution -SolutionId $solution.GetAttribute("id") -GacDeploy $GacDeploy -WebDeploy $WebDeploy -WebApp $WebApp -Force $Force
	}
	WaitDeployProcess -solutions $SolutionNodes -deploying $true
}

#
# Installs a solution from the solution store in the farm
#
function InstallSolution-Single($SolutionId, $WebApp, $Force, $XMLElement)
{
	[bool] $GacDeploy = [System.Convert]::ToBoolean($XMLElement.GetAttribute("GACdeploy"))
	[bool] $WebDeploy = [System.Convert]::ToBoolean($XMLElement.GetAttribute("webapp"))
	
	InstallSolution -SolutionId $SolutionId -GacDeploy $GacDeploy -WebDeploy $WebDeploy -WebApp $WebApp -Force $Force
	WaitDeployProcess-Single -solutionID $SolutionId -deploying $true
}

#
# Removes a solution from the solution store in the farm
#
function RemoveSolution($SolutionId, $Force) 
{
    Write-Host ("Removing solution " + $SolutionId)
	$SPsolution = Get-SPSolution -Identity $SolutionId -ErrorAction SilentlyContinue

	if ($SPsolution -eq $null)
	{
		Write-Host-Warning -message ("Solution " + $SolutionId + " is not present in the solution store")
	}
	else
	{
    	Remove-SPSolution -Identity $SolutionId -ErrorAction Continue -Force:$Force -Confirm:$false 
	}
}

#
# Removes some solutions from the solution store in the farm
#
function RemoveSolution-Multi($SolutionNodes, $Force) 
{
	foreach($solution in $SolutionNodes) 
	{
		RemoveSolution -SolutionId $solution.GetAttribute("id") -Force $Force -Confirm:$false 
	}
}

#
# Uninstalls a solution from the solution store in the farm
#
function UninstallSolution($SolutionId)
{
	Write-Host ("Uninstalling solution " + $SolutionId)
	$SPsolution = Get-SPSolution -Identity $SolutionId -ErrorAction SilentlyContinue 
	
	if ($SPsolution -eq $null)
	{
		Write-Host-Error -message ("Solution " + $SolutionId + " is not present in the solution store")
	}
	elseif (!$SPsolution.Deployed -and !$SPsolution.GlobalDeployed)
	{
		Write-Host-Warning -message ("Solution " + $SolutionId + " is not installed")
	}
	elseif ($SPsolution.ContainsWebApplicationResource)
	{
		Uninstall-SPSolution -Identity $SolutionId -AllWebApplications -Confirm:$false
	}
	else
	{
		Uninstall-SPSolution -Identity $SolutionId -Confirm:$false
	}
}

#
# Uninstall some solution from the solution store in the farm and waits the end of the process
#
function UninstallSolution-Multi($SolutionNodes)
{
	foreach( $solution in $SolutionNodes) 
	{
		UninstallSolution -SolutionId $solution.GetAttribute("id")
	}
	
	WaitDeployProcess -solutions $SolutionNodes -deploying $false 
}

#
# Uninstalls a single solution from the solution store in the farm and waits the end of the process
#
function UninstallSolution-Single($SolutionID)
{
	UninstallSolution -SolutionId $SolutionID
	WaitDeployProcess-Single -solutionID $SolutionID -deploying $false 
}

#
# Updates a solution in the solution store in the farm
#
function UpdateSolution($SolutionId, $SolutionPath, $Force) 
{
    Write-Host "Updating solution " -NoNewline
	Write-Host $SolutionId
	$SPsolution = Get-SPSolution -Identity $SolutionId -ErrorAction SilentlyContinue
	if ($SPsolution -eq $null)
	{
		Write-Host-Warning -message ("Solution " + $SolutionId + " is not present in the solution store")
	}
	else
	{
		Update-SPSolution –Identity $SolutionId -LiteralPath $SolutionPath -CASPolicies:$($SPsolution.ContainsCasPolicy) -GACDeployment:$($SPsolution.ContainsGlobalAssembly) -ErrorAction Continue -Force:$Force
	}
}

#
# Updates some solution in the solution store in the farm and waits the end of the process
#
function UpdateSolution-Multi($SolutionNodes, $RootPath, $Force) 
{
	foreach( $solution in $SolutionNodes) 
	{
		UpdateSolution -SolutionId $solution.GetAttribute("id") -SolutionPath ($RootPath + $solution.GetAttribute("id")) -Force $Force
	}
	WaitDeployProcess -solutions $SolutionNodes -deploying $true 
}

#
# Updates a single solution in the solution store in the farm and waits the end of the process
#
function UpdateSolution-Single($SolutionId, $SolutionPath, $Force) 
{
	UpdateSolution -SolutionId $SolutionId -SolutionPath $SolutionPath -Force $Force
	WaitDeployProcess-Single -solutionID $SolutionID -deploying $false 
}

#
# Prints the list of the solutions taken from the xml nodes
#
function PrintSolutions($SolutionNodes)
{
	$i = 1
	foreach( $solution in $SolutionNodes) 
	{
		$SPsolution = Get-SPSolution -Identity $solution.GetAttribute("id") -ErrorAction SilentlyContinue 
		
		$index = ([System.Convert]::ToString($i))
		$solutionName = FormatString -StringValue $solution.GetAttribute("id") -MaxLength 40 -Left $false -FillChar ' '
		Write-Host ($index + ") " + $solutionName + " ") -NoNewline

		if ($SPsolution -eq $null)
		{
			Write-Host "Not Existing" -ForegroundColor Red
		}
		elseif ($SPsolution.Deployed -or $SPsolution.GlobalDeployed)
		{
			Write-Host $SPsolution.DeploymentState -ForegroundColor Green
		}
		else
		{
			Write-Host $SPsolution.DeploymentState -ForegroundColor Yellow
		}
					
		$i = $i + 1
	}
}

#
# Prints the available list of features taken from the xml nodes
#
function PrintFeatures($FeatureNodes, $AppUrl)
{
	$i = 1
	foreach( $feature in $FeatureNodes) 
	{
		if ($feature.GetAttribute("scope") -eq "webapplication")
        {
            $targets = @($AppUrl)
        }
        else 
        {
            $targets = @($feature.targeturl)
        }    
        $statuses = @(GetFeature-Multi -FeatureId $feature.GetAttribute("id") -Scope $feature.GetAttribute("scope") -targets $targets)
       
        $featureNum = FormatString -StringValue ([System.Convert]::ToString($i)) -MaxLength 2 -Left $true -FillChar ' ' 
		$featureName = FormatString -StringValue $feature.GetAttribute("id") -MaxLength 50 -Left $false -FillChar ' '
		$scopeName = FormatString -StringValue ( "(" + $feature.GetAttribute("scope") +  ")") -MaxLength 10 -Left $false -FillChar ' '
        Write-Host ($featureNum + ") " + $featureName + " ") -NoNewline
        Write-Host $scopeName -ForegroundColor Yellow
            
        $k = 1
        foreach($target in $targets)
        {
    		Write-Host "  | "  -NoNewline

            if (($feature.GetAttribute("scope") -ne "allsites") -and ($feature.GetAttribute("scope") -ne "allwebs"))
            {
                $status = $statuses[($k - 1)];
        		if ($status -eq $false)
        		{
        			Write-Host "Not Activated" -ForegroundColor Red -NoNewline
        		}
        		else
        		{
        			Write-Host "Activated" -ForegroundColor Green -NoNewline
        		}
            }
            Write-Host ("    " +  $target) -ForegroundColor Yellow
            $k = $k + 1
		}
		$i = $i + 1
	}
}

#
# Prints the available list of features taken from the xml nodes
#
function PrintTargets($feature, $AppUrl)
{
	if ($feature.GetAttribute("scope") -eq "webapplication")
    {
        $targets = @($AppUrl)
    }
    else 
    {
        $targets = @($feature.targeturl)
    } 

	$statuses = @(GetFeature-Multi -FeatureId $feature.GetAttribute("id") -Scope $feature.GetAttribute("scope") -targets $targets)

	$k = 1
    foreach($target in $targets)
    {
		$featureNum = FormatString -StringValue ([System.Convert]::ToString($k)) -MaxLength 2 -Left $true -FillChar ' ' 
    	Write-Host ($featureNum + ") ") -NoNewline
		
        Write-Host $target -ForegroundColor Yellow -NoNewline

        if (($feature.GetAttribute("scope") -ne "allsites") -and ($feature.GetAttribute("scope") -ne "allwebs"))
        {
            $status = $statuses[($k - 1)];
        	if ($status -eq $false)
        	{
        		Write-Host "    Not Activated" -ForegroundColor Red
        	}
        	else
        	{
        		Write-Host "    Activated" -ForegroundColor Green
        	}
        }
        $k = $k + 1
	}
}

#
# Enables a feature based on ID and target URL
#
function EnableFeature($FeatureId, $TargetUrl)
{
    $Scope = Get-SPFeatureDefinitionScope $featureid
    $spfeaturedefinitionid = Get-SPFeatureDefinitionId $featureid
    
    if ($Scope -eq "site")
	{
        Enable-SPFeature-site -featureid $spfeaturedefinitionid -TargetUrl $targeturl
	}
	elseif ($Scope -eq "webapplication")
	{
		Enable-SPFeature -Identity $FeatureId -Url $TargetUrl -ErrorAction Continue
	}
	elseif ($Scope -eq "web")
	{
        Enable-SPFeature-web -featureid $spfeaturedefinitionid -TargetUrl $targeturl
	}
	else
	{
		Write-Host-Error "The scope for the feature was not found"
	}
}

#
# Disables a feature based on ID and target URL.
# Confirm controls whether confirmation is requested ($true) or not ($false)
#
function DisableFeature($FeatureId, $TargetUrl, $Confirm)
{
	$Scope = Get-SPFeatureDefinitionScope $featureid
    $spfeaturedefinitionid = Get-SPFeatureDefinitionId $featureid
    
    if ($Scope -eq "site")
	{
        Disable-SPFeature-site -featureid $spfeaturedefinitionid -TargetUrl $targeturl
	}
	elseif ($Scope -eq "webapplication")
	{
		Disable-SPFeature -Identity $FeatureId -Url $TargetUrl -Confirm:$Confirm -Force -ErrorAction Continue
	}
	elseif ($Scope -eq "web")
	{
        Disable-SPFeature-web -featureid $spfeaturedefinitionid -TargetUrl $targeturl
	}
}

function Enable-SPFeature-web($featureid, $TargetUrl)
{
    try {
        $web = get-spweb $targetUrl
        $web.features.add($featureid)
    }
    finally {
        if ($web -ne $null) {$web.dispose();}
        remove-variable web
    }
}

function Enable-SPFeature-site($featureid, $TargetUrl)
{
    try {
        $site = get-spsite $targetUrl
        $site.features.add($featureid)
    }
    finally {
        if ($site -ne $null) {$site.dispose();}
        remove-variable site
    }
}

function Disable-SPFeature-web($featureid, $TargetUrl)
{
    try {
        $web = get-spweb $targetUrl
        $web.features.remove($featureid)
    }
    finally {
        if ($web -ne $null) {$web.dispose();}
        remove-variable web
    }
}

function Disable-SPFeature-site($featureid, $TargetUrl)
{
    try {
        $site = get-spsite $targetUrl
        $site.features.remove($featureid)
    }
    finally {
        if ($site -ne $null) {$site.dispose();}
        remove-variable site
    }
}

function Get-SPFeatureDefinitionId($featureid)
{
    $spfeaturedefinition = Get-SPFeature -Identity $featureid -ErrorAction SilentlyContinue
    if ($spfeaturedefinition -eq $null) { return $null }
    return [System.GUID]($spfeaturedefinition.id.tostring())
}

function Get-SPFeatureDefinitionScope($featureid)
{
    $spfeaturedefinition = Get-SPFeature -Identity $featureid -ErrorAction SilentlyContinue
    if ($spfeaturedefinition -eq $null) { return $null }
    return $spfeaturedefinition.scope.tostring();
}

function Get-SPFeature-web($featureid, $TargetUrl)
{
    $spfeaturedefinitionid = Get-SPFeatureDefinitionId $featureid
    if ($spfeaturedefinitionid -eq $null) { return $false }
    
    $retval = $false
    try {
        $web = get-spweb $targetUrl
        
        $spfeature = $web.features[$spfeaturedefinitionid]
        if ($spfeature -eq $null)
        { $retval = $false }
        else
        { $retval = $true }
    }
    finally {
        if ($web -ne $null) {$web.dispose();}
        remove-variable web
    }
    
    return $retval;
}

function Get-SPFeature-site($featureid, $TargetUrl)
{
    $spfeaturedefinitionid = Get-SPFeatureDefinitionId $featureid
    if ($spfeaturedefinitionid -eq $null) { return $false }
    
    $site = get-spsite $targetUrl
    
    $spfeature = $site.features[$spfeaturedefinitionid]
    if ($spfeature -eq $null)
    { $retval = $false }
    else
    { $retval = $true }
    $site.dispose();
    remove-variable site
    
    return $retval;
}

function Get-SPFeature-webapp($featureid, $TargetUrl)
{
    $SPfeature = Get-SPFeature -Identity $FeatureId -WebApplication $TargetUrl -ErrorAction SilentlyContinue
    if ($spfeature -eq $null)
    { $retval = $false }
    else
    { $retval = $true }
    
    return $retval;
}

#
# Retrieve a feature
#
function GetFeature($FeatureId, $Scope, $TargetUrl) 
{
    $activated = $false;
    
    if ($Scope -eq "site")
	{
        $activated = Get-SPFeature-site -featureid $featureid -TargetUrl $TargetUrl
	}
	elseif ($Scope -eq "webapplication")
	{
		$activated = Get-SPFeature-webapp -featureid $featureid -TargetUrl $TargetUrl
	}
	elseif ($Scope -eq "web")
	{
        $activated = Get-SPFeature-web -featureid $featureid -TargetUrl $TargetUrl
	}
	return $activated;
}

#
# Retrieve a list of feature
#
function GetFeature-Multi($FeatureId, $Scope, $targets) 
{
    $statuses = new-object bool[] $targets.length;
    $i = 0;
	foreach($target in $targets)
    {
        $activated = GetFeature -FeatureId $FeatureId -Scope $Scope -TargetUrl $target
        $statuses[$i] = $activated;
        $i += 1;
    }
    return $statuses;
}

#
# Activate a feature
#
function ActivateFeature($FeatureId, $Scope, $TargetUrl, $Deactivate, $Confirm) 
{
    $retval = $null;
    
    if ($Deactivate -eq $true)
	{ $message = "Deactivating feature " }
    else
    { $message = "Activating feature " }
    
	$message += $FeatureId
    $message += (" [" + $TargetUrl + "]" )
    
    Write-Host ""
    Write-Host $message

    $activated = GetFeature -FeatureId $FeatureId -Scope $Scope -TargetUrl $TargetUrl
    
    if ($Deactivate -eq $true)
    {
        if ($activated -eq $false)
    	{
    		Write-Host-Warning ("Feature " + $FeatureId + " is not available in this scope (" + $Scope + ")")
    	}
    	else
    	{
            DisableFeature -FeatureId $FeatureId -TargetUrl $TargetUrl -Confirm $Confirm
            $activated = GetFeature -FeatureId $FeatureId -Scope $Scope -TargetUrl $TargetUrl
			if($activated -eq $false)
			{
				Write-Host-OK ("Feature " + $FeatureId + " has been deactivated successfully (" + $Scope + ")")
			}
			else
			{
				Write-Host-Error ("Feature " + $FeatureId + " has not been deactivated (" + $Scope + ")")
			}
    	}
    }
    else
    {
        if ($activated -eq $false)
    	{
    		EnableFeature -FeatureId $FeatureId -TargetUrl $TargetUrl
            $activated = GetFeature -FeatureId $FeatureId -Scope $Scope -TargetUrl $TargetUrl

			if($activated -eq $true)
			{
				Write-Host-OK ("Feature " + $FeatureId + " has been activated successfully (" + $Scope + ")")
			}
			else
			{
				Write-Host-Error ("Feature " + $FeatureId + " has not been activated (" + $Scope + ")")
			}
    	}
    	else
    	{
    		Write-Host-Warning -message ("Feature " + $FeatureId + " is already active")
    	}
    }
}

#
# Activate all targets for a feature
#
function ActivateFeatureAll($AppUrl, $FeatureId, $Scope, $targets, $TargetIndex, $Deactivate) 
{
    if ([int]($TargetIndex) -ge (0))
    {
        $targets = @($targets[$TargetIndex])
    }
    
    foreach($target in $targets)
    {
		if (($Scope -eq "allsites") -or ($Scope -eq "allwebs"))
        {
            ActivateFeatureCascade -AppUrl $AppUrl -FeatureId $FeatureId -Scope $Scope -target $target -Deactivate $Deactivate
        }
        else
        {
            ActivateFeature -FeatureId $FeatureId -Scope $Scope -TargetUrl $target -Deactivate $Deactivate -Confirm $true
        }
    }
}

#
# Activate feature on all available childs
#
function ActivateFeatureCascade($AppUrl, $FeatureId, $Scope, $target, $Deactivate)
{
    if ($Scope -eq "allsites")
    {
        Get-SPWebApplication $AppUrl | Get-SPSite -Limit ALL | ForEach-Object {
            ActivateFeature -FeatureId $FeatureId -Scope "site" -TargetUrl $_.Url -Deactivate $Deactivate -Confirm $false -OutVariable $OutVar -ErrorVariable $ErrorVar
            Write-Host $OutVar
            Write-Host $ErrorVar
        }
    }
    elseif ($Scope -eq "allwebs")
    {
        Get-SPWeb -site $target -Limit ALL | ForEach-Object {
            ActivateFeature -FeatureId $FeatureId -Scope "web" -TargetUrl $_.Url -Deactivate $Deactivate -Confirm $false -OutVariable $OutVar -ErrorVariable $ErrorVar
            Write-Host $OutVar
            Write-Host $ErrorVar
        }
    }    
}

#
# Activate multiple features
#
function ActivateFeatureAll-Multi($FeatureNodes, $AppUrl, $TargetIndex, $Deactivate) 
{
	foreach($Feature in $FeatureNodes) 
	{
        $FeatureId = $feature.GetAttribute("id");
        $scope = $feature.GetAttribute("scope");

        if ($scope -eq "webapplication")
        {
            $targets = @($AppUrl)
        }
        else 
        {
			$xPathExpr = "config/features/feature[@id='" + $Feature.id + "']/targeturl"
			$targets = @()
			foreach($target in $ConfigFile.SelectNodes($xPathExpr))
			{
				$targets += $target.InnerText
			}
        }
            
		Write-Host " Feature : " -NoNewline
		Write-Host $FeatureId -ForegroundColor Yellow
		Write-Host " Scope   : " -NoNewline
		Write-Host $scope -ForegroundColor Yellow
        
        if ((($scope -eq "allsites") -or ($scope -eq "allwebs")) -and !($NoGUI))
        {
            Write-Host "WARNING: This option will impact all the sites or webs in the targets."
            Write-Host "Are you REALLY sure to continue? (Y/N)"
    	    $OpConfirm = Read-Host
        	if(($OpConfirm -ne "Y") -and ($OpConfirm -ne "y"))
        	{
        		continue
        	}
        }

		ActivateFeatureAll -AppUrl $AppUrl -FeatureId $FeatureId -Scope $scope -targets $targets -TargetIndex $TargetIndex -Deactivate $Deactivate
	}
}

function OP-EnableDisableForceAttribute($ForceValue)
{
	$Force = $ForceValue
}

function GUI-EnableDisableForceAttribute()
{
	Clear-Host
	
	$ForceTitle = "Select the value of the force attribute:"
	$ForceChoices = @("1) True", "2) False")
			
	$StrForce = PrintChoiceMenu -Title $ForceTitle -Choices $ForceChoices
			
	if ($StrForce -eq "1")
		{OP-EnableDisableForceAttribute -Force $true}
	else
		{OP-EnableDisableForceAttribute -Force $false}

	PressEnterToContinue
}

function GUI-RestartOWSTimer()
{
	Clear-Host
    OP-RestartOWSTimer
	PressEnterToContinue
}

function OP-CollectFiles($fileType) 
{
	$xPathExpr = "config/solutions/solution"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)
		
	if($fileType -eq "wsp")
	{
		$WspPath = $RootPath + $ConfigFile.config.wsp_path 
		CollectFiles -destFolder $WspPath -reldbg $ConfName -ext "wsp" -SolutionNodes $solutions -srcPath $ConfigFile.config.solution_path.InnerText
	}
	elseif($fileType -eq "dll")
	{
		$DllPath = $RootPath + $ConfigFile.config.dll_path
		CollectFiles -destFolder $DllPath -reldbg $ConfName -ext "dll" -SolutionNodes $solutions -srcPath $ConfigFile.config.solution_path.InnerText
	}
}

function GUI-CollectWSPs() 
{
	Clear-Host
	OP-CollectFiles -fileType "wsp"
	PressEnterToContinue
}

function GUI-CollectDLLs() 
{
	Clear-Host
	OP-CollectFiles -fileType "dll"
	PressEnterToContinue
}

function OP-ListWebConfigModifications()
{
	Write-Host "Listing all WebConfig modifications."
	Write-Host "Web application: " -NoNewLine
	Write-Host $WebAppUrl -ForegroundColor Yellow
	Write-Host
	$configsToList = @()
	$config = $WebApp.WebConfigModifications
	foreach($c in $config) { $configsToList += $c }
	foreach($c in $configsToList) 
	{ 
		Write-Host ("        " + $c.Name)
	}
}

function GUI-ListWebConfigModifications()
{
	Clear-Host
	OP-ListWebConfigModifications
	PressEnterToContinue
}

function OP-CleanWebConfigModifications()
{
	Write-Host "Removing all WebConfig modifications."
	Write-Host "Web application: " -NoNewLine
	Write-Host $WebAppUrl -ForegroundColor Yellow
	Write-Host
	$configsToRemove = @()
	$config = $WebApp.WebConfigModifications
	foreach($c in $config) { $configsToRemove += $c }
	foreach($c in $configsToRemove) 
	{ 
		$RemoveResult = $WebApp.WebConfigModifications.Remove($c)

		if($RemoveResult -eq $true)
		{
			Write-Host-OK -message $c.Name
		}
		else
		{
			Write-Host-Error -message $c.Name
		}
	}
	$WebApp.Update()
	Write-Host -ForegroundColor Yellow "WebConfig modifications removed"
	Write-Host -ForegroundColor Yellow "Performing final clean"
	$WebApp.WebConfigModifications.clear()
	Write-Host -ForegroundColor Green "WebConfig modifications have been cleaned"
}

function GUI-CleanWebConfigModifications()
{
	Clear-Host
	OP-CleanWebConfigModifications
	PressEnterToContinue
}

function GUI-SelectSolution($solutions, $message)
{
	Clear-Host
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	Write-Host $message
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	PrintSolutions -SolutionNodes $solutions
	Write-Host "0) Back"
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
}

function GUI-SelectFeature($features, $message)
{
	Clear-Host
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	Write-Host $message
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	PrintFeatures -FeatureNodes $features -AppUrl $WebAppUrl
	Write-Host " 0) Back"
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
}

function OP-UninstallSingleSolution($SolutionID)
{
	$xPathExpr = ("config/solutions/solution[@id='" + $SolutionID + "']")
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	uninstallsolution-multi -SolutionNodes $solutions 
	removesolution-multi -SolutionNodes $solutions -Force $Force
}

function GUI-UninstallSingleSolution()
{
	$xPathExpr = "config/solutions/solution"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	$SolutionsNum = 0
	foreach($solution in $solutions) 
	{
		$SolutionsNum = $SolutionsNum + 1
	}

	while($true)
	{
		GUI-SelectSolution -solutions $solutions -message "Select the solution to uninstall"
		$SolChoice = Read-Host

		if ($SolChoice -eq "0")
		{
			break
		}
		elseif((Is-Int -Value $SolChoice) -and ([Int]$SolChoice -le $SolutionsNum))
		{
			$SolChoice -= 1
			$solution = $solutions.Item($SolChoice)
			
			Clear-Host
			OP-UninstallSingleSolution -SolutionID $solution.GetAttribute("id")
			PressEnterToContinue
			break
		}
		else
		{
			Clear-Host
			Write-Host "Unknown option. Please choose a different one."
			PressEnterToContinue
		}
	}
}

function OP-InstallSingleSolution($SolutionID)
{
	$WspPath = $RootPath + $ConfigFile.config.wsp_path 
	$WspPathConf = $WspPath + $ConfName + "\"

	$xPathExpr = "config/solutions/solution[@id='"+$SolutionID+"']"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	foreach($solution in $solutions)
	{
		AddSolution -SolutionId $SolutionID -SolutionPath ($WspPathConf + $SolutionID)
		InstallSolution-Single -SolutionID $SolutionID -WebApp $WebApp -Force $Force -XMLElement $solution
	}
}

function GUI-InstallSingleSolution()
{
	$xPathExpr = "config/solutions/solution"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	$SolutionsNum = 0
	foreach($solution in $solutions) 
	{
		$SolutionsNum = $SolutionsNum + 1
	}

	while($true)
	{
		GUI-SelectSolution -solutions $solutions -message "Select the solution to install"
		$SolChoice = Read-Host

		if ($SolChoice -eq "0")
		{
			break
		}
		elseif((Is-Int -Value $SolChoice) -and ([Int]$SolChoice -le $SolutionsNum))
		{
			$SolChoice -= 1
			$solution = $solutions.Item($SolChoice)
			
			Clear-Host
			OP-InstallSingleSolution -SolutionID $solution.GetAttribute("id")
			PressEnterToContinue
			break
		}
		else
		{
			Clear-Host
			Write-Host "Unknown option. Please choose a different one."
			PressEnterToContinue
		}
	}
}

function OP-UpdateSingleSolution($SolutionID)
{
	$WspPath = $RootPath + $ConfigFile.config.wsp_path 
	$WspPathConf = $WspPath + $ConfName + "\"

	UpdateSolution-Single -SolutionID $SolutionID -SolutionPath ($WspPathConf + $SolutionID) -Force $Force
}

function GUI-UpdateSingleSolution()
{
	$xPathExpr = "config/solutions/solution"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	$SolutionsNum = 0
	foreach($solution in $solutions) 
	{
		$SolutionsNum = $SolutionsNum + 1
	}

	while($true)
	{
		GUI-SelectSolution -solutions $solutions -message "Select the solution to update"
		$SolChoice = Read-Host

		if ($SolChoice -eq "0")
		{
			break
		}
		elseif((Is-Int -Value $SolChoice) -and ([Int]$SolChoice -le $SolutionsNum))
		{
			$SolChoice -= 1
			$solution = $solutions.Item($SolChoice)
			
			Clear-Host
			OP-UpdateSingleSolution -SolutionID $solution.GetAttribute("id")
			PressEnterToContinue
			break
		}
		else
		{
			Clear-Host
			Write-Host "Unknown option. Please choose a different one."
			PressEnterToContinue
		}
	}
}

function OP-UninstallAllSolutions()
{
	OP-UninstallFilteredSolutions -filter "todeploy"
}

function OP-UninstallFilteredSolutions($filter)
{
	$xPathExpr = "config/solutions/solution[@filter='"+$filter+"']"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	UninstallSolution-Multi -SolutionNodes $solutions 
	RemoveSolution-Multi -SolutionNodes $solutions -Force $Force
}

function GUI-UninstallAllSolutions()
{
	$ConfirmMultiTitle = "WARNING: this operation will uninstall all the solutions written in the config file. Are you sure?"
	$ConfirmMultiChoices = @("Y) Yes", "N) No")
			
	$StrConfirmMulti = PrintChoiceMenu -Title $ConfirmMultiTitle -Choices $ConfirmMultiChoices 
			
	if ($StrConfirmMulti -eq "Y")
		{$ConfirmMulti = $true}
	else
		{$ConfirmMulti = $false}

	Clear-Host
	if ($ConfirmMulti -eq $true)
	{
		OP-UninstallAllSolutions
		PressEnterToContinue
	}
}

function OP-InstallAllSolutions()
{
	OP-InstallFilteredSolutions -filter "todeploy"
}

function OP-InstallFilteredSolutions($filter)
{
	$xPathExpr = "config/solutions/solution[@filter='"+$filter+"']"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	$WspPath = $RootPath + $ConfigFile.config.wsp_path 
	$WspPathConf = $WspPath + $ConfName + "\"

	AddSolution-Multi -SolutionNodes $solutions -RootPath $WspPathConf
	InstallSolution-Multi -SolutionNodes $solutions -RootPath $WspPathConf -WebApp $WebApp -Force $Force
}

function GUI-InstallAllSolutions()
{
	$ConfirmMultiTitle = "WARNING: this operation will install all the solutions written in the config file. Are you sure?"
	$ConfirmMultiChoices = @("Y) Yes", "N) No")
			
	$StrConfirmMulti = PrintChoiceMenu -Title $ConfirmMultiTitle -Choices $ConfirmMultiChoices 
			
	if ($StrConfirmMulti -eq "Y")
		{$ConfirmMulti = $true}
	else
		{$ConfirmMulti = $false}

	Clear-Host
	if ($ConfirmMulti -eq $true)
	{
		OP-InstallAllSolutions
		PressEnterToContinue
	}
}

function OP-UpdateAllSolutions()
{
	OP-UpdateFilteredSolutions -filter "todeploy"
}

function OP-UpdateFilteredSolutions($filter)
{
	$xPathExpr = "config/solutions/solution[@filter='"+$filter+"']"
	$solutions = $ConfigFile.SelectNodes($xPathExpr)

	$WspPath = $RootPath + $ConfigFile.config.wsp_path 
	$WspPathConf = $WspPath + $ConfName + "\"

	UpdateSolution-Multi -SolutionNodes $solutions -RootPath $WspPathConf -Force $Force
}

function GUI-UpdateAllSolutions()
{
	$ConfirmMultiTitle = "WARNING: this operation will update all the solutions written in the config file. Are you sure?"
	$ConfirmMultiChoices = @("Y) Yes", "N) No")
			
	$StrConfirmMulti = PrintChoiceMenu -Title $ConfirmMultiTitle -Choices $ConfirmMultiChoices 
			
	if ($StrConfirmMulti -eq "Y")
		{$ConfirmMulti = $true}
	else
		{$ConfirmMulti = $false}

	Clear-Host
	if ($ConfirmMulti -eq $true)
	{
		OP-UpdateAllSolutions
		PressEnterToContinue
	}
}

function GUI-PrintTargets($feature, $targets)
{
	Clear-Host
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	Write-Host "Select one of the available targets: [Press 'Enter' to select all targets]"
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
	PrintTargets -feature $feature -TargetNodes $targets -AppUrl $WebAppUrl
	Write-Host " 0) Back"
	Write-Host (FormatString -StringValue "" -MaxLength 79 -Left $false -FillChar '-')
}

function OP-EnableDisableSingleFeature($FeatureNodes, $TargetIndex, $Deactivate)
{
	ActivateFeatureAll-Multi -FeatureNodes $FeatureNodes -AppUrl $WebAppURL -TargetIndex $TargetIndex -Deactivate $Deactivate
}

function GUI-SelectTarget($feature, $deactivate)
{
	$xPathExpr = "config/features/feature[@id='" + $feature.id + "']/targeturl"
	$targets = $ConfigFile.SelectNodes($xPathExpr)

	$TargetNum = 0
	foreach($target in $targets) 
	{
		$TargetNum = $TargetNum + 1
	}

	while($true)
	{
		GUI-PrintTargets -feature $feature -targets $targets
		$Choice = Read-Host

		if ($Choice -eq "0")
		{
			break
		}
		elseif((Is-Int -Value $Choice) -and ([Int]$Choice -le $TargetNum))
		{
			$Choice -= 1;			
			Clear-Host
			OP-EnableDisableSingleFeature -FeatureNodes $feature -TargetIndex $Choice -Deactivate $deactivate
			PressEnterToContinue			
			break
		}
		else
		{
			Clear-Host
			Write-Host "Unknown option. Please choose a different one."
			PressEnterToContinue
		}
	}
}

function GUI-EnableDisableSingleFeature($Deactivate)
{
	Clear-Host
	$xPathExpr = "config/features/feature[@id]"
	$features = $ConfigFile.SelectNodes($xPathExpr)

	$FeaturesNum = 0
	foreach($feature in $features) 
	{
		$FeaturesNum = $FeaturesNum + 1
	}
	
	while($true)
	{
		$message = "Select the feature to "
		if($Deactivate -eq $true)
		{
			$message += "deactivate"
		}
		else
		{
			$message += "activate"
		}

		GUI-SelectFeature -features $features -message $message

		$Choice = Read-Host

		if ($Choice -eq "0")
		{
			break
		}
		elseif((Is-Int -Value $Choice) -and ([Int]$Choice -le $FeaturesNum))
		{
			$Choice = ($Choice - 1)
			$feature = $features.Item($Choice)

			Clear-Host
			GUI-SelectTarget -feature $feature -Deactivate $Deactivate
			break
		}
		else
		{
			Clear-Host
			Write-Host "Unknown option. Please choose a different one."
			PressEnterToContinue
		}
	}
}

function GUI-DisableSingleFeature()
{
	GUI-EnableDisableSingleFeature -Deactivate $true
}

function GUI-EnableSingleFeature()
{
	GUI-EnableDisableSingleFeature -Deactivate $false
}

function OP-EnableDisableAllFeatures($Deactivate)
{
	$xPathExpr = "config/features/feature"
	$features = $ConfigFile.SelectNodes($xPathExpr)

	$WspPath = $RootPath + $ConfigFile.config.wsp_path 
	$WspPathConf = $WspPath + $ConfName + "\"

	ActivateFeatureAll-Multi -FeatureNodes $features -AppUrl $WebAppURL -TargetIndex -1 -Deactivate $Deactivate
}

function OP-DisableAllFeatures()
{
	OP-EnableDisableAllFeatures -Deactivate $true
}

function OP-EnableAllFeatures()
{
	OP-EnableDisableAllFeatures -Deactivate $false
}

function GUI-DisableAllFeatures()
{
	$ConfirmMultiTitle = "WARNING: this operation will disable all the features written in the config file. Are you sure?"
	$ConfirmMultiChoices = @("Y) Yes", "N) No")
			
	$StrConfirmMulti = PrintChoiceMenu -Title $ConfirmMultiTitle -Choices $ConfirmMultiChoices 
			
	if ($StrConfirmMulti -eq "Y")
		{$ConfirmMulti = $true}
	else
		{$ConfirmMulti = $false}

	Clear-Host
	if ($ConfirmMulti -eq $true)
	{
		OP-DisableAllFeatures
		PressEnterToContinue
	}
}

function GUI-EnableAllFeatures()
{
	$ConfirmMultiTitle = "WARNING: this operation will enable all the features written in the config file. Are you sure?"
	$ConfirmMultiChoices = @("Y) Yes", "N) No")
			
	$StrConfirmMulti = PrintChoiceMenu -Title $ConfirmMultiTitle -Choices $ConfirmMultiChoices 
			
	if ($StrConfirmMulti -eq "Y")
		{$ConfirmMulti = $true}
	else
		{$ConfirmMulti = $false}

	Clear-Host
	if ($ConfirmMulti -eq $true)
	{
		OP-EnableAllFeatures
		PressEnterToContinue
	}
}

function GUI()
{
	do
	{
		Clear-Host

		$MainTitle = "Select one of the available options:"

		$MainChoices = @(
			" 1) Enable/Disable force attribute",
			" 2) Restart OwsTimer",
			" 3) Deactivate all features",
			" 4) Deactivate single feature",			
			" 5) Collect WSPs",
			" 6) Collect DLLs",			
			" 7) List WebConfig modifications",
			" 8) Clean WebConfig modifications",
			" 9) Uninstall single solution",
			"10) Uninstall all solutions",
			"11) Install single solution",
			"12) Install all solutions",
			"13) Update single solution",
			"14) Update all solutions",
			"15) Activate all features",
			"16) Activate single feature",
			"0 ) Exit")
		
		$OpChoice = PrintChoiceMenu -Title $MainTitle -Choices $MainChoices

		if ($OpChoice -eq "0"){ break }
		elseif($OpChoice -eq "1"){ GUI-EnableDisableForceAttribute } 
		elseif($OpChoice -eq "2"){ GUI-RestartOWSTimer }
		elseif($OpChoice -eq "3"){ GUI-DisableAllFeatures }
		elseif($OpChoice -eq "4"){ GUI-DisableSingleFeature }
		elseif($OpChoice -eq "5"){ GUI-CollectWSPs }
		elseif($OpChoice -eq "6"){ GUI-CollectDLLs }
		elseif($OpChoice -eq "7"){ GUI-ListWebConfigModifications }
		elseif($OpChoice -eq "8"){ GUI-CleanWebConfigModifications }
		elseif($OpChoice -eq "9"){ GUI-UninstallSingleSolution }
		elseif($OpChoice -eq "10"){ GUI-UninstallAllSolutions }
		elseif($OpChoice -eq "11"){ GUI-InstallSingleSolution }
		elseif($OpChoice -eq "12"){ GUI-InstallAllSolutions }
		elseif($OpChoice -eq "13"){ GUI-UpdateSingleSolution }
		elseif($OpChoice -eq "14"){ GUI-UpdateAllSolutions }
		elseif($OpChoice -eq "15"){ GUI-EnableAllFeatures }
		elseif($OpChoice -eq "16"){ GUI-EnableSingleFeature }

		else
		{
			Clear-Host
			Write-Host "Unknown or inactive option. Please choose a different one."
			PressEnterToContinue
			continue
		}

	} while (1)
}

function MessageAndContinue($Message, $MessageType)
{
	Write-Host
	Write-host 
	Write-host 
	Write-host " ============================================================================"
	Write-Host
	
	if($MessageType -eq "Error"){ Write-Host-Error -Message $Message }
	elseif($MessageType -eq "Warning"){ Write-Host-Warning -Message $Message }
	elseif($MessageType -eq "OK"){ Write-Host-OK -Message $Message }
	else{ Write-Host $Message -ForegroundColor Yellow }

	Write-Host
	Write-host " ============================================================================"
	Write-Host
	Write-Host
	Write-Host
	PressEnterToContinue
}


function AskContinue($Message)
{
	MessageAndContinue -Message $Message -MessageType "Warning"
}

function No-GUi
{
	$xPathExpr = "config/steps/step"
	$steps = $ConfigFile.SelectNodes($xPathExpr)
	$stepNum = 1
	$stepTot = 0;
	$stepSkip = 0;

	$elapsedTotTimer = [Diagnostics.Stopwatch]::StartNew()
	$elapsedLastStepTimer = [Diagnostics.Stopwatch]::StartNew()

	foreach($step in $steps)
	{
		$stepTot += 1
	}

	foreach($step in $steps)
	{
		if($stepNum -lt $stepSkip) 
		{ 
			$stepNum += 1
			continue 
		} 

		$elapsedLastStepTimer.Stop()
		$elapsedLastStep = $elapsedLastStepTimer.Elapsed
		$elapsedLastStepMessage = ("Elapsed from last: " + [string]$elapsedLastStep.Hours + ":" + [string]$elapsedLastStep.Minutes + ":" + [string]$elapsedLastStep.Seconds + "." + [string]$elapsedLastStep.Milliseconds )
		$elapsedLastStepTimer = [Diagnostics.Stopwatch]::StartNew()

		$elapsedTot = $elapsedTotTimer.Elapsed
		$elapsedTotMessage = ("Elapsed from start: " + [string]$elapsedTot.Hours + ":" + [string]$elapsedTot.Minutes + ":" + [string]$elapsedTot.Seconds + "." + [string]$elapsedTot.Milliseconds )
		
		Write-host 
		Write-host 
		Write-host " ============================================================================"
		Write-host 
		Write-host (" STEP " + $stepNum + "    " + $elapsedLastStepMessage + "    " + $elapsedTotMessage)
		Write-host 
		Write-host " ============================================================================"
		Write-host 
		Write-host 

		if($step.op -ne $null)
		{
			if($step.op -eq "NoOp") 
			{
				if($step.message -ne $null)
				{
					Write-Host (" " + $step.message)
				}
			} 
			elseif($step.op -eq "EnableForceAttribute") { OP-EnableDisableForceAttribute -Force $true }
			elseif($step.op -eq "DisableForceAttribute") { OP-EnableDisableForceAttribute -Force $false }
			elseif($step.op -eq "CollectWSPs") { OP-CollectFiles -fileType "wsp" }
			elseif($step.op -eq "CollectDLLs") { OP-CollectFiles -fileType "dll" }
			elseif($step.op -eq "RestartOWSTimer") { OP-RestartOWSTimer }
			elseif($step.op -eq "ListWebConfigModifications") { OP-ListWebConfigModifications }
			elseif($step.op -eq "CleanWebConfigModifications") { OP-CleanWebConfigModifications }
			elseif($step.op -eq "DisableAllFeatures") { OP-DisableAllFeatures }
			elseif(($step.op -eq "DisableSingleFeature") -or ($step.op -eq "EnableSingleFeature")) 
			{ 
				$Deactivate = $true
				if($step.op -eq "EnableSingleFeature") { $Deactivate = $false }

				if($step.id -ne $null)
				{
					$xPathExprStep = "config/features/feature[@id='" + $step.id + "']"
					$FeatureNodes = $ConfigFile.SelectNodes($xPathExprStep)

					$Choice = [int](-1)
					OP-EnableDisableSingleFeature -FeatureNodes $FeatureNodes -TargetIndex $Choice -Deactivate $Deactivate
				}
				else
				{
					Write-Host-Error -Message "<id> tag is missing"
				}				
			}
			elseif($step.op -eq "EnableAllFeatures") { OP-EnableAllFeatures }
			elseif($step.op -eq "UninstallAllSolutions") { OP-UninstallAllSolutions }
			elseif($step.op -eq "UninstallFilteredSolutions") 
			{ 
				if($step.filter -ne $null) { OP-UninstallFilteredSolutions -filter $step.filter }
				else { Write-Host-Error -Message "<filter> tag is missing" }				
			}
			elseif($step.op -eq "UninstallSingleSolution") 
			{ 
				if($step.id -ne $null) { OP-UninstallSingleSolution -SolutionID $step.id }
				else { Write-Host-Error -Message "<id> tag is missing" }				
			}
			elseif($step.op -eq "InstallAllSolutions") { OP-InstallAllSolutions }
			elseif($step.op -eq "InstallFilteredSolutions") 
			{ 
				if($step.filter -ne $null) { OP-InstallFilteredSolutions -filter $step.filter }
				else { Write-Host-Error -Message "<filter> tag is missing" }				
			}
			elseif($step.op -eq "InstallSingleSolution") 
			{ 
				if($step.id -ne $null) { OP-InstallSingleSolution -SolutionID $step.id }
				else { Write-Host-Error -Message "<id> tag is missing" }				
			}
			elseif($step.op -eq "UpdateAllSolutions") { OP-UpdateAllSolutions }
			elseif($step.op -eq "UpdateFilteredSolutions") 
			{ 
				if($step.id -ne $null) { OP-UpdateFilteredSolutions -filter $step.filter }
				else { Write-Host-Error -Message "<filter> tag is missing" }				
			}
			elseif($step.op -eq "UpdateSingleSolution") 
			{ 
				if($step.id -ne $null) { OP-UpdateSingleSolution -SolutionID $step.id }
				else { Write-Host-Error -Message "<id> tag is missing" }				
			}			
			else { Write-Host-Warning -Message "<op> contains an unrecognized command" }			
		}

		if($step.continue -ne $null)
		{
			AskContinue -Message $step.continue
		}

		if($step.gui -ne $null)
		{
			MessageAndContinue -Message $step.gui
			GUI
			Clear-Host
		}

		if($step.exit -ne $null)
		{
			$Title = (" " + $step.exit)
			$Choices = @("Y) Yes - Exit", "N) No - Continue")
			$OpChoice = $null

			Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
			Write-Host $Title -ForegroundColor Red
			do
			{
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
				for ($i = 0; $i -lt $Choices.Length; $i++)
				{
					Write-Host (" " + $Choices[$i])
				}
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
				$OpChoice = Read-Host 
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
			}while(($OpChoice -ne "Y") -and ($OpChoice -ne "N")) 
						
			if ($OpChoice -eq "Y")
			{
				break
			}
		}

		if($step.skip -ne $null)
		{
			$Title = (" " + $step.skip)
			$OpChoice = $null

			Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
			Write-Host $Title -ForegroundColor Green
			do
			{
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
				Write-Host ("Input a number between " +  ($stepNum + 1) + " and " + $stepTot + " to advance to that step.")
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
				$OpChoice = Read-Host 
				Write-Host (" " + (FormatString -StringValue "" -MaxLength 78 -Left $false -FillChar '-'))
			}while(!(Is-Int -Value $OpChoice) -or ([int]$OpChoice -le $stepNum) -or ([int]$OpChoice -gt $stepTot))
						
			$stepSkip = [int]($OpChoice)
		}

		$stepNum += 1
	}

	$elapsedLastStepTimer.Stop()
	$elapsedLastStep = $elapsedLastStepTimer.Elapsed
	$elapsedLastStepMessage = ("Elapsed from last: " + [string]$elapsedLastStep.Hours + ":" + [string]$elapsedLastStep.Minutes + ":" + [string]$elapsedLastStep.Seconds + "." + [string]$elapsedLastStep.Milliseconds )

	$elapsedTotTimer.Stop()
	$elapsedTot = $elapsedTotTimer.Elapsed
	$elapsedTotMessage = ("Elapsed from start: " + [string]$elapsedTot.Hours + ":" + [string]$elapsedTot.Minutes + ":" + [string]$elapsedTot.Seconds + "." + [string]$elapsedTot.Milliseconds )
		
	Write-host 
	Write-host
	Write-host " ============================================================================"
	Write-host 
	Write-host (" FINISHED!    " + $elapsedLastStepMessage + "    " + $elapsedTotMessage)
	Write-host 
	Write-host " ============================================================================"
	Write-host 
	Write-host
}

function Main($RootPath)
{
	# In GUI mode, display the splash screen
	GUI-SplashScreen			

	# 1. Loading configuration file.
	$ConfigPath = $RootPath + "\config.xml"
	Write-host " ---------------------------------------------------"
	Write-host 
	Write-host " Configuration file:        " -NoNewline
	Write-Host $ConfigPath -ForegroundColor Yellow

	if(($LogFilePath -ne $null) -and ($LogFilePath -ne ""))
	{
		Write-Host " Log file path:             " -NoNewline
		Write-Host ($LogFilePath) -ForegroundColor Yellow
	}

	[Xml]$ConfigFile = Get-Content $ConfigPath -ErrorAction SilentlyContinue
	if ($ConfigFile -eq $null)
	{
		Write-Host -Fore Red "Unable to read config file"
	    Exit 1
	}

	write-host " Target environment:        " -NoNewline
	write-host $ConfigFile.config.environment -ForegroundColor Yellow

	$WebAppUrl = $ConfigFile.config.WebAppURL
	$ConfName = $ConfigFile.config.configuration
	
	Write-Host " Target application URL:    "-NoNewline
	Write-host $WebAppURL -ForegroundColor Yellow
	Write-Host " Build configuration:       "-NoNewline
	Write-host $ConfName -ForegroundColor Yellow

	Write-Host " Force attribute:           "-NoNewline
	Write-host "Disabled" -ForegroundColor Red
	Write-Host " Run date:                  " -NoNewline
	Write-host (Get-Date) -ForegroundColor Green

	Write-host 
	Write-host " ---------------------------------------------------"
	Write-host 

	$WebApp = Get-SPWebApplication -Identity $WebAppURL -ErrorAction SilentlyContinue	
	if ($WebApp -eq $null)
	{
		Write-Host " Web Application $WebAppURL not found. Aborting." -ForegroundColor Red
		Exit 1
	}
	else
	{
		if(!$NoGUI)
		{
			PressEnterToContinue
		}
	}

	if($NoGUI)
	{
		No-GUI
	}
	else
	{
		GUI
		Clear-Host
	}
}

#
# Call the main entry point
#

$logPath = "transcript.txt"
if(($LogFilePath -ne $null) -and ($LogFilePath -ne ""))
{
	$logPath = $LogFilePath
}

Start-Transcript $logPath
main -RootPath ($myinvocation.mycommand.path | Split-Path)
Stop-Transcript
#
# This cmdlet takes care of the disposable objects to prevent memory leak.
#
Stop-SPAssignment -Global
