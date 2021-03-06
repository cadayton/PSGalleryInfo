# PSGalleryInfo
Display and search for top downloaded modules or scripts from the PowerShell Gallery.
By default, MicroSoft Corporation and PowerShell DSC author are excluded.

Output is displayed in a Out-GridView and selection of an entry will display the module's
or script's project web page, if one exists.

Module and script data is cached locally for a period of 8 hours, so repeated execution of
the command is very quick.

The script extends the work of Chris Hunt.
https://www.automatedops.com/blog/2017/04/28/bringing-the-community-forward/

**Install from PowerShell Gallery**

    Install-Script PSGalleryInfo -Scope currentuser

**Example**

    PSGalleryInfo

Displays the top 100 downloaded modules from the PowerShell Gallery, PSGallery.

For other examples.

    Get-Help PSGalleryInfo -full
