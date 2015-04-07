# New-VsScProj
Create New Visual Studio Sitecore Project that can be used to target Sitecore Azure Website Services.

# Description #
This PowerShell script automates the creation of initial  Visual Studio (VS) project for Sitecore (SC) development.  Below are list of activities that the script automates:

1.  Copy existing VS template folder to target destination
2.  Copy/Extract SC files into target destination
3.  Update target VS project file (name, assembly, namespace, content, references) 
4.  Copy contents of Sitecore/Data folder into VS/App_Data folder
5.  Update sc.variable for "dataFolder" in web.config to /App_Data.  
6.  Add empty readme.txt file every folder that does not have any file or folders.  This is done to insure the folders persists when they are deployed to Azure platform.

 
# Usage #
Below outlines the script usage and parameters.  Note that after running the script to create the Visual Studio Project, you must add the license.xml file into the data folder (e.g. App_Data).  


    .\New-VsScProj.ps1 -source <p1>              `
                       [-vsTemplateFolder <p2>]  `
                       [-vsTargetFolder   <p3>]  `
                       [-overwriteExistingFiles <p3>]

**-source** parameter is *required*, and must point to the full path of Sitecore installation zip file (manual) or respective extracted folder.  The root Sitecore distribution folder must contain folders: Data, Database, and Website.  If a path to a zip file is specified, then the script first extracts the file into the current user's temporary folder, and then it deletes the temporary folder at the end of the script.

**-vsTemplateFolder** parameter is *optional*.  A Sitecore Visual Studio project is created by first copying an existing VS project.  If omitted, the script uses the first VS folder that contains the word "Template" within the PowerShell script path as the template source parameter.  As part of this distribution, there exists an empty MVC/WebForms VS project folder "TemplateWebFormsMVC" within the PowerShell script folder.

**-vsTargetFolder** parameter is *optional*. It is target folder for the VS project folder.  If omitted, then the current working folder is used for the target folder.  Note that the PowerShell script enforces the following conventions: *1)* VS project file (.csproj) will match the target folder name, *2)* If no target folder is specified, then the vsTemplateFolder name is used for the target folder name (excluding the word template found in the vsTemplateFolder), *3)* if no target folder is specified and there already exists a target folder, then a new folder with a numeric suffix is created.  For example, if WebFormsMvc folder already exists, then WebFormsMVC1 is created.

**-overwriteExistingFiles** parameter is optional (default false).  If target folder is specified via -vsTargetFolder and if the target already exists, then this flag is responsible for updating existing files and project.  By existing files are not replaced.

# Examples #
    New-ScProj -source 'c:\Sitecore8.zip' 

    New-ScProj -source 'c:\Sitecore8.zip' `
               -vsTargetFolder "c:\temp\sc8"

    New-ScProj -source 'c:\Sitecore8.zip'       `
               -vsTemplateFolder 'c:\Sitecore8' `
               -vsTargetFolder "c:\temp\sc8"

    New-ScProj -source 'c:\Sitecore8.zip'       `
               -vsTemplateFolder 'c:\Sitecore8' `
               -vsTargetFolder "c:\temp\sc8"    `
               -overwriteExistingFiles $true

