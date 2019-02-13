#PowerShell: GeneralEnvironmentDeploymentWithCsv.ps1
################################
########## PARAMETERS ##########
################################ 
[CmdletBinding()]
Param(
    # SsisServer is required
    [Parameter(Mandatory=$True,Position=1)]
    [string]$SsisServer,
     
    # CatalogFolderName is required 
    [Parameter(Mandatory=$True,Position=2)]
    [string]$CatalogFolderName,
     
    # EnvironmentName is required
    [Parameter(Mandatory=$True,Position=3)]
    [string]$EnvironmentName,
     
    # FilepathCsv is required
    [Parameter(Mandatory=$True,Position=4)]
    [string]$FilepathCsv,

    # IsPacFilePath is required
    [Parameter(Mandatory=$True,Position=5)]
    [string]$IspacFilePath,

    # ProjectName is not required
    # If empty filename is used
    [Parameter(Mandatory=$False,Position=6)]
    [string]$ProjectName,

    # Use environment Variable
    [Parameter(Mandatory=$True,Position=7)]
    [string]$UseEnvVariables
)
 
clear
Write-Host "========================================================================================================================================================"
Write-Host "==                                                                 Used parameters                                                                    =="
Write-Host "========================================================================================================================================================"
Write-Host "SSIS Server             :" $SsisServer
Write-Host "Environment Name        :" $EnvironmentName
Write-Host "Environment Folder Path :" $CatalogFolderName
Write-Host "Filepath of CSV file    :" $FilepathCsv
Write-Host "IspacFile Path          :" $IspacFilePath
Write-Host "Project Name            :" $ProjectName
Write-Host "========================================================================================================================================================"
 
 
#########################
########## CSV ##########
#########################
# Check if ispac file exists
if (-Not (Test-Path $FilepathCsv))
{
    Throw  [System.IO.FileNotFoundException] "CSV file $FilepathCsv doesn't exists!"
}
else
{
    $FileNameCsv = split-path $FilepathCsv -leaf
    Write-Host "CSV file" $FileNameCsv "found"
}

###########################
########## ISPAC ##########
###########################
# Check if ispac file exists
if (-Not (Test-Path $IspacFilePath))
{
    Throw  [System.IO.FileNotFoundException] "Ispac file $IspacFilePath doesn't exists!"
}
else
{
    $IspacFileName = split-path $IspacFilePath -leaf
    Write-Host "Ispac file" $IspacFileName "found"
}
 
 
############################
########## SERVER ##########
############################
# Load the Integration Services Assembly
Write-Host "Connecting to SSIS server $SsisServer "
$SsisNamespace = "Microsoft.SqlServer.Management.IntegrationServices"
[System.Reflection.Assembly]::LoadWithPartialName($SsisNamespace) | Out-Null;
 
# Create a connection to the server
$SqlConnectionstring = "Data Source=" + $SsisServer + ";Initial Catalog=master;Integrated Security=SSPI;"
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection $SqlConnectionstring
 
# Create the Integration Services object
$IntegrationServices = New-Object $SsisNamespace".IntegrationServices" $SqlConnection
 
# Check if connection succeeded
if (-not $IntegrationServices)
{
  Throw  [System.Exception] "Failed to connect to SSIS server $SsisServer "
}
else
{
   Write-Host "Connected to SSIS server" $SsisServer
}
 
 
#############################
########## CATALOG ##########
#############################
# Create object for SSISDB Catalog
$Catalog = $IntegrationServices.Catalogs["SSISDB"]
 
# Check if the SSISDB Catalog exists
if (-not $Catalog)
{
    # Catalog doesn't exists. The user should create it manually.
    # It is possible to create it, but that shouldn't be part of
    # deployment of packages or environments.
    Throw  [System.Exception] "SSISDB catalog doesn't exist. Create it manually!"
}
else
{
    Write-Host "Catalog SSISDB found"
}
 
 
############################
########## FOLDER ##########
############################
# Create object to the (new) folder
$Folder = $Catalog.Folders[$CatalogFolderName]
 
# Check if folder already exists
if (-not $Folder)
{
    # Folder doesn't exists, so create the new folder.
    Write-Host "Creating new folder" $CatalogFolderName
    $Folder = New-Object $SsisNamespace".CatalogFolder" ($Catalog, $CatalogFolderName, $CatalogFolderName)
    $Folder.Create()
}
else
{
    Write-Host "Folder" $CatalogFolderName "found"
}

 

#############################
########## PROJECT ##########
#############################
# Deploying project to folder
if($Folder.Projects.Contains($ProjectName)) {
    Write-Host "Deploying" $ProjectName "to" $FolderName "(REPLACE)"
}
else
{
    Write-Host "Deploying" $ProjectName "to" $FolderName "(NEW)"
}
# Reading ispac file as binary
[byte[]] $IspacFile = [System.IO.File]::ReadAllBytes($IspacFilePath)
$Folder.DeployProject($ProjectName, $IspacFile)
$Project = $Folder.Projects[$ProjectName]
if (-not $Project)
{
    # Something went wrong with the deployment
    # Don't continue with the rest of the script
    return ""
}



######### Do not use Environment Varibales######
if ($UseEnvVariables -eq 0)
{
Import-CSV $FilepathCsv -Header Datatype,ParameterName,ParameterValue,ParameterDescription,Sensitive -Delimiter ';' | Foreach-Object{
 If (-not($_.Datatype -eq "Datatype"))
 {
  if ($project.Parameters.Contains($_.ParameterName))  
  {
   Write-Host $_.ParameterName "project parameter found"
   $project.Parameters[$_.ParameterName].Set([Microsoft.SqlServer.Management.IntegrationServices.ParameterInfo+ParameterValueType]::Literal,$_.ParameterValue)
   $Project.Alter()
   Write-Host $_.ParameterName "project parameter value updated to " $_.ParameterValue
  }
 }
}
###########################
########## READY ##########
###########################
# Kill connection to SSIS
$IntegrationServices = $null
Write-Host "Project Deployment Finished"
}

else
{ 
#################################
########## ENVIRONMENT ##########
#################################
# Create object for the (new) environment
$Environment = $Catalog.Folders[$CatalogFolderName].Environments[$EnvironmentName]
 
# Check if folder already exists
if (-not $Environment)
{
    Write-Host "environment" $EnvironmentName in $CatalogFolderName 
    $Environment = New-Object $SsisNamespace".EnvironmentInfo" ($Folder, $EnvironmentName, $EnvironmentName)
    $Environment.Create()
}
else
{
    Write-Host "Environment" $EnvironmentName "found with" $Environment.Variables.Count "existing variables"
    # Optional: Recreate to delete all variables, but be careful:
    # This could be harmful for existing references between vars and pars
    # if a used variable is deleted and not recreated.
     $Environment.Drop()
     $Environment = New-Object $SsisNamespace".EnvironmentInfo" ($folder, $EnvironmentName, $EnvironmentName)
     $Environment.Create()
}


#################################
######ENVIRONMENT REFERNCE#######
#################################
# Create object for the environment
$Environment = $Catalog.Folders[$CatalogFolderName].Environments[$EnvironmentName]
 
if ($Project.References.Contains($EnvironmentName, $CatalogFolderName))
{
    Write-Host "Reference to" $EnvironmentName "found and getting dropped and re-created"
    $project.References.Remove($EnvironmentName, $folder.Name)
    $project.References.Add($EnvironmentName, $folder.Name)
    $project.Alter() 
}
else
{
     Write-Host "Adding reference to" $EnvironmentName
     $Project.References.Add($EnvironmentName, $CatalogFolderName)
     $Project.Alter() 
} 
 
###############################
########## VARIABLES ##########
###############################
$InsertCount = 0
$UpdateCount = 0
 
 
Import-CSV $FilepathCsv -Header Datatype,ParameterName,ParameterValue,ParameterDescription,Sensitive -Delimiter ';' | Foreach-Object{
 If (-not($_.Datatype -eq "Datatype"))
 {
  #Write-Host $_.Datatype "|" $_.ParameterName "|" $_.ParameterValue "|" $_.ParameterDescription "|" $_.Sensitive
  # Get variablename from array and try to find it in the environment
  $Variable = $Catalog.Folders[$CatalogFolderName].Environments[$EnvironmentName].Variables[$_.ParameterName]
 
 
  # Check if the variable exists
  if (-not $Variable)
  {
   # Insert new variable
   Write-Host "Variable" $_.ParameterName "added"
   $Environment.Variables.Add($_.ParameterName, $_.Datatype, $_.ParameterValue, [System.Convert]::ToBoolean($_.Sensitive), $_.ParameterDescription)
 
   $InsertCount = $InsertCount + 1
  }
  else
  {
   # Update existing variable
   Write-Host "Variable" $_.ParameterName "updated"
   $Variable.Type = $_.Datatype
   $Variable.Value = $_.ParameterValue
   $Variable.Description = $_.ParameterDescription
   $Variable.Sensitive = [System.Convert]::ToBoolean($_.Sensitive)
 
   $UpdateCount = $UpdateCount + 1
  }
 }
} 
$Environment.Alter()


########################################
########## PROJECT PARAMETERS ##########
########################################
$ParameterCount = 0
# Loop through all project parameters
foreach ($Parameter in $Project.Parameters)
{
    # Get parameter name and check if it exists in the environment
    $ParameterName = $Parameter.Name
    if ($ParameterName.StartsWith("INTERN_","CurrentCultureIgnoreCase"))
    {
        # Internal parameters are ignored (where name starts with INTERN_)
        Write-Host "Ignoring Project parameter" $ParameterName " (internal use only)"
    }
    elseif ($Environment.Variables.Contains($Parameter.Name))
    {
        $ParameterCount = $ParameterCount + 1
        Write-Host "Project parameter" $ParameterName "connected to environment"
        $Project.Parameters[$Parameter.Name].Set([Microsoft.SqlServer.Management.IntegrationServices.ParameterInfo+ParameterValueType]::Referenced, $Parameter.Name)
        $Project.Alter()
    }
    else
    {
        # Variable with the name of the project parameter is not found in the environment
        # Throw an exeception or remove next line to ignore parameter
        "Project parameter $ParameterName doesn't exist in environment"
    }
}
Write-Host "Number of project parameters mapped:" $ParameterCount

 
 
###########################
########## READY ##########
###########################
# Kill connection to SSIS
$IntegrationServices = $null
Write-Host "Project deployment Finished, total inserts" $InsertCount " and total updates" $UpdateCount
}