<#PSScriptInfo

  .VERSION 1.0.0
  .GUID 0fd916fe-3a0d-48c4-a196-18ea085e071f
  .AUTHOR Craig Dayton
  .COMPANYNAME Example.com
  .COPYRIGHT Absolute Zero
  .TAGS 
  .LICENSEURI 
  .PROJECTURI https://github.com/cadayton/PSGalleryInfo
  .ICONURI 
  .EXTERNALMODULEDEPENDENCIES 
  .REQUIREDSCRIPTS 
  .EXTERNALSCRIPTDEPENDENCIES 
  .RELEASENOTES

#>

<#
  .SYNOPSIS
    Display and search for top downloaded modules or scripts in PowerShell Gallery, PSGallery
    By default, MicroSoft Corporation and PowerShell DSC authors are excluded.

  .DESCRIPTION
    Display and search for top downloaded modules or scripts from the PowerShell Gallery, PSGallery
    By default, MicroSoft Corporation and PowerShell DSC author are excluded.

    Optionally, one can specific a different registered repository name.

    Output is displayed in a Out-GridView and selection of an entry will display the module's
    or script's project web page, if one exists.

    Module and script data is cached locally for a period of 8 hours, so repeated execution of
    the command is very quick.

    The script extends the work of Chris Hunt.
    https://www.automatedops.com/blog/2017/04/28/bringing-the-community-forward/

  .PARAMETER scripts
    Specify this option to display top X downloaded scripts

  .PARAMETER top
    Default is 100.

  .PARAMETER all
    Switch to include Microsoft code.

  .PARAMETER matchAuthor
    String value to display specific modules or scripts matching the Author Name property.

  .PARAMETER matchDesc
    String value to display specific modules or scripts matching the Description property.

  .PARAMETER repository
    String value to specific a specific registered PowerShell gallery.  The default PowerShell is
    PSGallery.

  .INPUTS
    ScriptRepo.xml or ModuleRepo.xml depending on age of the file.

  .OUTPUTS
    ScriptRepo.xml or ModuleRepo.xml depending on parameters specified.
    Out-Gridview

  .EXAMPLE
    PSGalleryInfo

    Displays the top 100 downloaded modules minus default exclusions.

    On the Out-GridView dialog select an entry and click on the "OK" button
    to display the module's project home page.  Or click on the "Cancel" button
    to exit.

  .EXAMPLE
    PSGalleryInfo -All

    Displays the top 100 downloaded modules from PSGallery without any exclusions.

    On the Out-GridView dialog select an entry and click on the "OK" button
    to display the module's project home page.  Or click on the "Cancel" button
    to exit.

  .EXAMPLE
    PSGalleryInfo -repository MyGallery

    Same as the prior example, but specifies a private PowerShell repository named, MyGallery.

    See the following URL on how to create your own internal PowerShell repository.
    https://kevinmarquette.github.io/2017-05-30-Powershell-your-first-PSScript-repository/?utm_source=rss&utm_medium=blog&utm_content=rss

  .EXAMPLE
    PSGalleryInfo -script -all

    Displays the top 100 downloaded scripts without any exclusions.

  .EXAMPLE
    PSGalleryInfo -matchAuthor "Lee"

    Displays the top 100 downloaded modules with an author property matching "Lee".

  .NOTES
    Author: Craig Dayton
      1.0.0: 05/01/2017 - Extending the work of Chris Hunt
        https://www.automatedops.com/blog/2017/04/28/bringing-the-community-forward/
    
#>

# PSGalleryInfo Params
  [cmdletbinding()]
    Param(
      [Parameter(Position=0,
        Mandatory=$false,
        HelpMessage = "Search for PowerShell Gallery scripts",
        ValueFromPipeline=$True)]
        [switch]$script,
      [Parameter(Position=1,
        Mandatory=$false,
        HelpMessage = "Display Top X modules or scripts",
        ValueFromPipeline=$True)]
        [ValidateNotNullorEmpty()]
        [int]$top = 100,
      [Parameter(Position=2,
        Mandatory=$false,
        HelpMessage = "Display Online Lottery Results & update history file",
        ValueFromPipeline=$True)]
        [ValidateNotNullorEmpty()]
        [switch]$all,
      [Parameter(Position=3,
        Mandatory=$false,
        HelpMessage = "Display all game history file records",
        ValueFromPipeline=$True)]
        [ValidateNotNullorEmpty()]
        [string]$matchAuthor,
      [Parameter(Position=4,
        Mandatory=$false,
        HelpMessage = "Display the pick history of all games",
        ValueFromPipeline=$True)]
        [ValidateNotNullorEmpty()]
        [string]$matchDesc,
      [Parameter(Position=5,
        Mandatory=$false,
        HelpMessage = "Specify a repository name",
        ValueFromPipeline=$True)]
        [ValidateNotNullorEmpty()]
        [string]$repository = "PSGallery"
   )
#

# Declarations

  #Import-Module BurntToast;
  $sPath          = Get-Location;
  $ScriptRepoFile = "$sPath\ScriptRepo-$repository.xml";
  $ModuleRepoFile = "$sPath\ModuleRepo-$repository.xml";

#

# functions

  function Get-XMLFile {
    param ([string]$XMLInput)

    [int]$mts = 0;

    if (Test-Path $XMLInput) {
      $csvFile = Get-ChildItem $XMLInput;
      $cts = $csvFile.LastWriteTime
      $nts = New-TimeSpan -Start (Get-Date) -End $cts
      # Number of minutes since LastWriteTime
      [int]$mts = (($nts.Days * 1440) + ($nts.Hours * 60) + ($nts.Minutes)) * -1;
    }

    if (($mts -gt 480) -or ($mts -eq 0)) { # older than 8 hours or file doesn't exist
      return $null
    } else {
      $obj = Import-CliXML -Path $XMLInput
      return $obj
    }
  }
#

# Main Routine

  if ($script) {
      $codetype = "scripts";
      $modules = Get-XMLFile $ScriptRepoFile
      if ($modules -is [Object]) { } else {
        Write-Progress -Activity "Loading script data from PowerShell Gallery $repository..." -Status "Please wait"
        $modules = Find-Script -Repository $repository
        Write-Progress -Activity "Done" -Completed;
        $modules | Export-CliXML -Path $ScriptRepoFile
      }   
  } else {
      $codetype = "modules"
      $modules = Get-XMLFile $ModuleRepoFile
      if ($modules -is [Object]) { } else {
        Write-Progress -Activity "Loading module data from PowerShell Gallery $repository..." -Status "Please wait"
        $modules = Find-Module -Repository $repository
        Write-Progress -Activity "Done" -Completed;
        $modules | Export-CliXML -Path $ModuleRepoFile
      }
  }

  if ((-not $matchAuthor) -and (-not $matchDesc)) {
    $normalized = $modules |
      Select-Object Name, @{Name = "Downloads"; Expression = {$_.AdditionalMetadata.downloadCount -as [int]}}, Author, Description, ProjectUri
  }

  if ($matchAuthor) {
    $normalized = $modules |
    Where-Object {$_.Author -match $matchAuthor} |
    Select-Object Name, @{Name = "Downloads"; Expression = {$_.AdditionalMetadata.downloadCount -as [int]}}, Author, Description, ProjectUri
  }

  if ($matchDesc) {
    $normalized = $modules |
    Where-Object {$_.Description -match $matchDesc} |
    Select-Object Name, @{Name = "Downloads"; Expression = {$_.AdditionalMetadata.downloadCount -as [int]}}, Author, Description, ProjectUri
  }

  if (-not $All) {
      $Selected = $normalized |
          Where-Object {$_.Author -ne 'Microsoft Corporation' -and $_.Author -ne 'PowerShell DSC'} |
          Sort-Object Downloads -Descending |
          Select-Object Name, Downloads, Author, Description, ProjectUri -first $top | 
          Out-GridView -Title "Top $top Non-Microsoft $codetype" -OutputMode Single;
  } else {
      $Selected = $normalized |
          Sort-Object Downloads -Descending |
          Select-Object Name, Downloads, Author, Description, ProjectUri -first $top | 
          Out-GridView -Title "Top 100 $codetype" -OutputMode Single;
  }

  if ($Selected -eq $null) {
    Write-Host " Selection Cancelled" -ForegroundColor Red
    Write-Host "Thanks for trying! Bye" -ForegroundColor Blue
  } else {
    [string]$URL = $Selected.ProjectUri;
    [string]$PRJ = $Selected.Name;
    if ($URL -ne $null -and $URL.Length -gt 7) {
      $Browser = new-object -com internetexplorer.application;
      $Browser.navigate2($URL);
      $Browser.visible = $true;
    } else {
      Write-Host "$PRJ " -NoNewLine -ForegroundColor Blue
      Write-Host "doesn't reference a home page" -ForegroundColor Red
    }
  }
#