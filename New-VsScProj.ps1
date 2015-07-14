################################################################################
################################################################################
# Author:  Edvin Eshagh
#
# Date:    3/19/2015
#
# Purpose: Create a visual studio project that is inclusive of all Sitecore
#          files so that SC can be deployed to target environment.
#          The project came into inception because I wanted to deploy SC 
#          onto Azure Website service, and I needed to create new project
#          files quickly and easily. 
#
# Usage:   
#   New-VsScProj -source <p1>           `
#                -vsTemplateFolder <p2> `
#                -vsTargetFolder <p3>   `
#                -overwriteExistingFiles $false
#
#       -soruce (required):
#            Sitecore root folder that contains Data, Website, and 
#            Database folders.  The sourceFolder may also be 
#            Sitecore distribution zip file. 
#
#       -vsTemplateFolder (optional):
#            visual studio project that is to be used as a 
#            template for creating a new project that contains the 
#            sitecore files. If none specified, then the powershell
#            script folder is searched for a template folder.
#
#       -vsTargetFolder (optional):
#            target destination for new visual studio project.  
#            As convention, the leaf folder name is used as the 
#            project name, assembly and namespace.
#
#       -overwriteExistingFiles (optional - default false):
#            If set to true, then existing files in the target folder 
#            are overwriten during the copy process.  Otherwise,
#            the will not be altered.
#   
################################################################################
################################################################################

Param(
    [Parameter(Mandatory=$false, HelpMessage=
    "Sitecore zip file or extracted root folder")]
    [string] $source,

    [Parameter(Mandatory=$false, HelpMessage=
    "Visual studio folder that serves as a template for a new project")]
    [string] $vsTemplateFolder,
    
    [Parameter(Mandatory=$false, HelpMessage=
    "Target path to new visual studio project folder. " +
    "If target project exists, then it is updated " +
    "(as apose to using the template).  If no " +
    "path is specified, then the current working directory is used")]
    [string] $vsTargetFolder,
    
    [Parameter(Mandatory=$false, HelpMessage=
    "Overwrite existing files")]
    [string] $overwriteExistingFiles = $false
)

Set-StrictMode –Version 2
$ErrorActionPreference = "Stop"

################################################################################
# programmer note:
#     declaring [xml] $xml=get-content(...) as shown here:
#     https://www.simple-talk.com/sysadmin/powershell/powershell-data-basics-xml
#     does not preserve sitecore variables in web.config when $xml.save() 
#     method is called.  Therefore, .NET object is used to load and save:
#       $xml = new-object System.Xml.XmlDcoument
#        $xml.load(...)
#        ...
#        $xml.save(...)
################################################################################


################################################################################
# CONSTANTS
#
if (! (Test-Path -Path Variable:VS_TEMPLATE_REGEX)) {

    #####################################
    # VS Template projects within the
    # script folder must contain
    # the following string so that 
    # the user does not need to specify
    # a VS template folder path
    #
    Set-Variable VS_TEMPLATE_REGEX -option Constant -value "Template"
    
    #####################################
    # Azure does not preserve empty 
    # folders.  To preserve the folder, 
    # an empty file must be added.
    # When set to true, an empty file 
    # is added to each empty folder.
    #
    Set-Variable ADD_EMPTY_FILE_TO_EMPTY_FOLDER -Option Constant -Value $true

    #####################################    
    # The following variable defines the
    # empty file name to be added to 
    # empty folders
    #
    Set-Variable EMPTY_FILENAME_STR -Option Constant -Value "readme.txt"

    #####################################    
    # Typescript settings
    #
    Set-Variable TYPE_SCRIPT_VERSION -Option Constant -Value "1.0"
    Set-Variable TYPE_SCRIPT_MODULE -Option Constant -Value "amd"
        
    #####################################
    # File extensions used to determine
    # the XML node name added to VS
    # project file
    #
    Set-Variable FILE_EXT_TO_VS_NODE_MAP -Option Constant -Value @{
        ".ts"     = "TypeScriptCompile";
        ".resx"   = "EmbeddedResource";
        "default" = "Content";
        "folder"  = "Folder";
    }
}



################################################################################
# Get-VsProjectFilePath()
#
# Purpose: Find the visual studio project file (*.csproj)
#
# Param:   $folder - full path to a folder that should contain 
#                    VS project file
#   
# Return:  Full path to the csproj file within the specified folder.
#          It is assumed that the project file matches the folder name.
#          If no match is found, then the first *.csproj file is returned.
#
function Get-VsProjectFilePath ($folder) {

    #####################################
    # $param validation
    #
    
    # if the parameter is a file, return null if it isn't the project file
    if (Test-Path -PathType leaf -Path $folder) {
        if ($folder.ToLower().EndsWith(".csproj") -eq $true) {
            return $folder
        }
        return $null
    }
    # return $null if the target $folder does not exist
    else {
        if (!(Test-Path -PathType container -Path $folder)) {
            return $null
        }
    }
    
    #######################################
    # Find project file that matches folder
    #
    $local:projFiles= Get-ChildItem -Path $folder |
        Where-Object{$_.Extension -eq ".csproj"} 
    
    $local:expectedProjName = $folder.SubString($folder.LastIndexOf("\")+1)
    
    $local:projFile = ($projFiles | 
        Where-Object {$_.BaseName -match $expectedProjName} | 
        Select-Object -First 1)
    
    #######################################
    # Use the first project file if no
    # project file matches folder name
    #
    if (! $projFile) {
        $projFile = ($projFiles | Select-Object -First 1)
    }
    
    if ($projFile -ne $null) {
        return $projFile.FullName
    }
    else {
        return $null
    }
    
} # Get-VsProjectFilePath()


################################################################################
# Copy-VsTemplateFolder()
#
# Purpose: Copy visual studio project folder
#          
# Param:  $srcFolder - Source folder that contains *.csproj file
#
#         $dstFolder - destination folder for duplicating the $srcFolder
#
# Return: If the following conditions exists, then false is returned:
#              $srcFolder does not contain a VS project file (e.g. *.csproj)
#           or $dstFolder already contains a VS project file (e.g. *.csproj)
#
#           Otherwise $true is return after performing the following actions:
#              1) $srcFolder is copied to $dstFolder and 
#              2) the destination project file name, project namespace, 
#                 and project assembly are modified to match the 
#                 destination folder name
#
function Copy-VsTemplateFolder ($srcFolder, $dstFolder) {

    #####################################
    # $param validation 
    #
    $local:srcVsProjFile = (Get-VsProjectFilePath $srcFolder).replace("/","\\")
    
    New-Item -Path $dstFolder -ItemType directory `
            -ErrorAction Ignore | Out-Null
            
    $local:dstVsProjFile = (Get-VsProjectFilePath $dstFolder)
    
    if ($srcVsProjFile -eq $null -or $dstVsProjFile -ne $null) {
        return $false
    }
    
    $srcFolder = $srcVsProjFile.SubString(0, $srcVsProjFile.LastIndexOf("\"))
    
    $local:srcProjName = Get-ChildItem -Path $srcVsProjFile | 
        Select -First 1 -ExpandProperty BaseName

    # strip out the suffix slash
    $dstFolder = $dstFolder -replace "[\\/]\s*$", ""
    
    $local:dstProjName = $dstFolder.substring($dstFolder.lastIndexOf("\") +1)
    
    #####################################
    # Copy source to destination
    #
    
    # Create destination folder if it desn't exist
    if (! (Test-Path -PathType Container -Path $dstFolder)) {
    
        New-Item -Path $dstFolder -ItemType directory `
            -ErrorAction Ignore | Out-Null
    }
    
     Get-ChildItem -Path $srcFolder | ForEach-Object {
    
        Copy-Item $_.FullName $dstFolder -recurse `
        -ErrorAction SilentlyContinue 
    }
    
    $projFiles = Get-ChildItem -Path $dstFolder | 
        Where-Object{    
            $_.Extension -eq ".csproj" -or $_.Extension -eq ".user"
        } 
    
    #####################################
    # Rename destination project files    
    #
    $srcVsProjFile = "$dstFolder\$srcProjName.csproj"
    $dstVsProjFile = "$dstFolder\$dstProjName.csproj"
    
    Rename-Item $srcVsProjFile $dstVsProjFile
    
    Rename-Item "$srcVsProjFile.user" "$dstVsProjFile.user" `
        -ErrorAction Ignore

    #####################################
    # Update VS project assembly/namespace
    #
    $local:xml = new-object System.Xml.XmlDocument
    $xml.Load($dstVsProjFile)
    
    $local:nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsMgr.AddNamespace("ns", $xml.DocumentElement.NamespaceURI)
    
    $local:propertyGroup = (
        Get-XmlNode -XmlDocument $xml `
            -NodePath "/Project/PropertyGroup/RootNamespace" `
            -NodeSeparatorCharacter "/" `
            ).ParentNode
        
    #$local:propertyGroup = $xml.Project.PropertyGroup | Where-Object {$_.RootNamespace -ne $null}

    $propertyGroup.RootNamespace       `
       = $propertyGroup.AssemblyName   `
       = $dstProjName
       
    $xml.Save($dstVsProjFile)
    
    return $true
    
} # Copy-VsTemplateFolder()


################################################################################
# Copy-FilesIntoProject()
#
# Purpose: Copy files from source to destination, and then add them
#          to the visual studio project file.
#
# Param:
#         $vsTargetProjFilePath - full path to the visual studio *.csproj file
#
#         $srcFolder - folder path to copy files from
#
#         $inclusionMatchFilter - regular expression to test against full
#                                 file path for inclusion.  If blank or null
#                                 then all files are included.
#
#         $exclusionMathFilter - regular expression to test against full 
#                                 file path for exclusion.  If blank or null
#                                 then "no" file is excluded
#
#         $vsLogicalFolder - target folder within visual studio where files
#                            should be added (e.g. /bin, /app_data, etc)
#
#         $overwriteExistingFiles - If true, then files are overwritten during
#                            copy process; otherwise, existing files are not
#                            modified.
#
# Return: N/A
#
function Copy-FilesIntoProject (
            $vsTargetProjFilePath, 
            $srcFolder, 
            $inclusionMatchFilter, 
            $exclusionMatchFilter,
            $vsLogicalFolder, 
            $overwriteExistingFiles) {
    
    #####################################
    # Parameter cleansing
    #
    
    # strip-out slashes from begining and end of folder 
    # prefix slashes maybe needed for \\UNC source path
    $srcFolder = $srcFolder -replace "[\\/]\s*$", ""  
    
    $vsLogicalFolder  = $vsLogicalFolder  `
        -replace "(^\s*[\\/]*)|([\\/]\s*$)", ""
    
    # xml Content/Folder node within 
    # VS project *.csproj file
    $local:node = $null 
    
    $local:vsLogicalRelativePath = $null
    
    $local:vsProjFolder = $vsTargetProjFilePath.
        SubString(0, $vsTargetProjFilePath.LastIndexOf("\") )

    #####################################
    # Load VS *.csproj as XML
    #
    $local:xml = new-object System.Xml.XmlDocument
    $xml.Load($vsTargetProjFilePath)

    $local:itemGroup = $xml.CreateElement(
        "ItemGroup", $xml.Project.xmlns)    
        
    $xml.Project.AppendChild($itemGroup) | Out-Null
    

    #####################################
    # Build a hash table of all the 
    # content items within visual studio, 
    # so that we don't have to do a 
    # sequential search for their 
    # existance.  We don't want to re-add
    # something that is already there.
    #
    $local:includedFiles = @{}     # All files and folders 
                                   # already within visual studio
                                   
    $local:foldersHaveFiles = @{}  # Used to find empty folders
    
    $local:itemGroupChildNodes = (
        Get-XmlNodes -XmlDocument $xml `
        -NodePath ".Project.ItemGroup/child::node()" `
        -NodeSeparatorCharacter ".")
        
    # populate $includedFiles
    
    $itemGroupChildNodes | 
        Where-Object { 
            $_.GetAttribute("Include").Length -gt 0                           `
            -and $FILE_EXT_TO_VS_NODE_MAP.Values -contains $_.Get_LocalName() `
        } |
        ForEach-Object {
            $local:contentPath = $_.GetAttribute("Include").ToLower()
            $includedFiles.Set_Item($contentPath, $_)
            
            # Add parent folder if it does not exist
            $local:parentPath = Split-Path -Parent $contentPath.ToLower()
            if ($parentPath.length -gt 0) {
                $includedFiles.Set_Item($parentPath, $_)
            }
        }

    #####################################
    # Check every file from $srcFolder 
    # to see if it satisfies the filter 
    # condition
    #
    Get-ChildItem -Path $srcFolder -Recurse |
        
        # apply inclusion & exclusion filter
        Where-Object { ([string]::IsNullOrWhiteSpace($inclusionMatchFilter) `
                       -or $_.FullName -match $inclusionMatchFilter)        `
                       -and                                                 `
                       ([string]::IsNullOrWhiteSpace($exclusionMatchFilter) `
                       -or $_.FullName -iNotMatch $exclusionMatchFilter)
        } |
        
        #####################################    
        # Copy files from source to dest.
        # and then insert each file into VS
        #
        ForEach-Object {
        
            $local:srcFile = $_
            
            $vsLogicalRelativePath = $vsLogicalFolder +
                $_.FullName.substring($srcFolder.length)
                
            $vsLogicalRelativePath = 
                $vsLogicalRelativePath -replace "^\s*[\\/]",""
    
            #$vsLogicalRelativePath # write-host

            $local:dstFilePath = $vsProjFolder + "\" + 
                $vsLogicalFolder + 
                $_.FullName.substring($srcFolder.length)
                
            $dstFilePath = $dstFilePath.replace("\\", "\")
   
            #####################################
            # create empty folder so copy-item 
            # does not fail for non-existing 
            # folders
            #
            if ($srcFile -is [System.IO.DirectoryInfo]) {
            
                New-Item -Path $dstFilePath -ItemType directory `
                    -ErrorAction Ignore | Out-Null
            }
            
            #####################################
            # copy file to target
            #
            else { 
                if ($overwriteExistingFiles -eq $true) {
                    write-host "Copy-Item" $srcFile.FullName $dstFilePath 
                    Copy-Item $srcFile.FullName $dstFilePath -Force 
                    
                } else {
                
                    if (!(Test-Path -PathType Leaf -Path $dstFilePath )) {
                        write-host "Copy-Item" $srcFile.FullName $dstFilePath 
                        Copy-Item $srcFile.FullName $dstFilePath 
                        #-ErrorAction SilentlyContinue
                    }
                    else {
                        Write-Host "Skiping" $srcFile.FullName
                    }
                }
                
            }
            
            #####################################
            # Update VS project file for the 
            # newly copied file if its not there
            # already.
            #
            if (! $includedFiles.ContainsKey(
                $vsLogicalRelativePath.ToLower())) {
            
                # update the hash so we don't
                # do a sequential search
                #
                $includedFiles.Set_Item(
                    $vsLogicalRelativePath.ToLower(), $true)

                #####################################
                # Do not add folders to VS project 
                # because later in this module,
                # we will add them if they don't 
                # already have any files.
                #
                if ($srcFile -is [System.IO.DirectoryInfo]) {
                
                    $foldersHaveFiles.Set_Item(
                        $vsLogicalRelativePath.ToLower(), $false)
                }
                
                #####################################
                # Add file to VS project file based
                # on its extension because diffent
                # file types should be added with
                # different xml nodes.  For example,
                # CSS files should be <Content...
                # and resource files should be 
                # <EmbeddedResource ... and so on.
                #
                else {            
                    
                    $local:fileExtension = ""
                    if ($vsLogicalRelativePath.LastIndexOf(".") -gt 0) {
                    
                        $fileExtension = $vsLogicalRelativePath.
                            ToLower().
                            SubString(
                            $vsLogicalRelativePath.LastIndexOf("."))
                            
                    }
                    
                    $local:nodeName = $null
                                    
                    $nodeName = $FILE_EXT_TO_VS_NODE_MAP.
                        Get_Item("default")
                    
                    if ($FILE_EXT_TO_VS_NODE_MAP.
                        ContainsKey($fileExtension)) {
                        
                        $nodeName = $FILE_EXT_TO_VS_NODE_MAP.
                            Get_Item($fileExtension)
                    }
                    
                    $node = $xml.
                        CreateElement($nodeName, $xml.Project.xmlns)
                        
                    $node.setAttribute("Include", $vsLogicalRelativePath)
                    
                    $ItemGroup.AppendChild($node) | Out-Null
                }
                
                #####################################
                # Update local hash variable to keep
                # track of empty folders 
                #
                $local:isVsRootFile = 
                    $vsLogicalRelativePath.LastIndexOf("\") -lt 1
                    
                if (! $isVsRootFile) {
                
                    $local:folder = $vsLogicalRelativePath.ToLower().
                        SubString(0,$vsLogicalRelativePath.LastIndexOf("\"))
                        
                    $foldersHaveFiles.Set_Item($folder, $true)
                }
                
            } # if (! $includedFiles.ContainsKey(...
            
        } # ForEach-Object

    #####################################
    # For empty folders, we must manually 
    # add them (since we ignored them 
    # earlier).  Depending on global 
    # constant, we may add an empty 
    # readme.txt file instead of an
    # empty folder
    # 
    $foldersHaveFiles.GetEnumerator()      | 
        Where-Object {$_.value -eq $false} |
        ForEach {
            $vsLogicalRelativePath = ($_.key -replace "^\s*[\\/]","")
            
            if ($ADD_EMPTY_FILE_TO_EMPTY_FOLDER `
            -and ![string]::IsNullOrWhiteSpace($EMPTY_FILENAME_STR)) {
            
                $vsLogicalRelativePath = 
                    $vsLogicalRelativePath + "\$EMPTY_FILENAME_STR"
                    
                $local:txtFile = 
                    ($vsProjFolder + "\" + $vsLogicalRelativePath).
                    replace("\\", "\")
                    
                New-Item $txtFile -ItemType file `
                    -ErrorAction Ignore| Out-Null
                
                $node = $xml.CreateElement("Content", $xml.Project.xmlns)
            }
            else {
                $node = $xml.CreateElement("Folder", $xml.Project.xmlns)
            }

            $vsLogicalRelativePath # Write-Host            
            
            # update project if the path doesn't exist
            if (! $includedFiles.ContainsKey(
                $vsLogicalRelativePath.ToLower())) {

                $includedFiles.Set_Item(
                    $vsLogicalRelativePath.ToLower(), $true)
                    
                $node.setAttribute("Include", $vsLogicalRelativePath)
                $ItemGroup.AppendChild($node)    | Out-Null
            }
            
        }

    # Update visual studio if we changed it.
    if ($ItemGroup.ChildNodes.count -gt 0) {
        $xml.Save($vsTargetProjFilePath.ToString())
        
    }
    
} # Copy-FilesIntoProject()


################################################################################
# Add-TypeScriptSupportToVs()
#
# Purpose: Updates visual studio project file to have TypeScript support 
#            with Asynchronous Module Definition (amd) support. 
#
# Param:   $vsProjFilePath - Visual studio file to update
#
# Result:  If the specified file does not exist or if it already has 
#          TypeScript definition, then no action is taken.
#          Otherwise, the *.csproj file is modified by adding the 
#          following xml nodes to the project file:
<#
                <project>
                    <Import Project="$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.Default.props" Condition="Exists('$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.Default.props')" />
                    <Import Project="$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.targets" Condition="Exists('$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.targets')" />
                    <PropertyGroup Condition="...">
                          <TypeScriptToolsVersion>1.0</TypeScriptToolsVersion>
                        <TypeScriptModuleKind>amd</TypeScriptModuleKind>
                    </PropertyGroup>
                 ...
                </project>
#>
#
function Add-TypeScriptSupportToVs($vsProjFilePath) {

	#####################################
    # Param validation
    #
    if (! (Test-Path -PathType Leaf -Path $vsProjFilePath)) {
        return
    }
    
    #####################################
    # load XML
    #
    $local:xml = new-object System.Xml.XmlDocument
    
    $xml.Load($vsProjFilePath)
    
    $local:hasTypeScript = ($xml.Project.Import |
        Where-Object {$_.Project -imatch "TypeScript"})
    
    if ($hasTypeScript -eq $true) {
        return 
    }
    
    $local:node = $null

    #####################################
    # Mutate XML
    #

    # <Import Project 1
    $node = $xml.CreateElement("Import", $xml.Project.xmlns)
    
    $node.SetAttribute("Project", "`$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v`$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.Default.props")
    
    $node.SetAttribute("Condition", "Exists('`$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v`$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.Default.props')")
    
    $xml.Project.InsertBefore($node, $xml.Project.ChildNodes[0]) | Out-Null
    
    $local:propertyGroup = (
        Get-XmlNode -XmlDocument $xml `
            -NodePath "/Project/PropertyGroup" `
            -NodeSeparatorCharacter "/" `
            )
    
    # PropertyGroup.TypeScriptToolsVersion
    $propertyGroup.AppendChild( 
        $xml.CreateElement("TypeScriptToolsVersion", 
        $xml.Project.xmlns)).innerText = $TYPE_SCRIPT_VERSION

    #PropertyGroup.TypeScriptModuleKind
    $propertyGroup.AppendChild( 
        $xml.CreateElement("TypeScriptModuleKind", 
        $xml.Project.xmlns)).innerText = $TYPE_SCRIPT_MODULE


    #####################################
    # <Import Project 2
    # Location of the Import is important, 
    # which must come after specifying
    # the TypeScript Module compile type.
    # Otherwise, we'll get the this error
    #    cannot compile external modules unless the '--module' flag is provided
    #
    # reference: http://stackoverflow.com/questions/25147727/typescript-external-module
    #
    $node = $xml.CreateElement("Import", $xml.Project.xmlns)
    
    $node.SetAttribute("Project", "`$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v`$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.targets")
    
    $node.SetAttribute("Condition", "Exists('`$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v`$(VisualStudioVersion)\TypeScript\Microsoft.TypeScript.targets')")
    
    $xml.Project.InsertAfter($node, $propertyGroup) | Out-Null
    
            
    #####################################
    # Save XML
    #
    $xml.Save($vsProjFilePath)

} # Add-TypeScriptSupportToVs()


################################################################################
# Add-EmptyFolderIntoProject()
#
# Purpose: Add an empty folder to the file system and visual studio project.
#          The intend of this method is to add a folder that contains
#          a readme.txt file, so that the folder persists on Azure website.
#          In absence of any files, empty folders are deleted from Azure website.
#
# Param:   $vsTargetProjFilePath - Visual studio project file (*.csproj) where
#               a folder must be added to.
#
#          $vsLogicalFolder - The logical path to the folder within VS.
#               For example: /temp
#
# Result:  If the global flag $ADD_EMPTY_FILE_TO_EMPTY_FOLDER is set to $true, 
#          then an empty readme.txt file is added into the folder and the 
#          respective file is added into visual studio project.  However, if the
#          flag is $false, then only the empty folder is added to the VS project
#
function Add-EmptyFolderIntoProject(
            $vsTargetProjFilePath, 
            $vsLogicalFolder) {
    
    #####################################
    # Param validation & cleansing
    
    # strip-out slashes from begining and end of folder 
    # prefix slashes maybe needed for \\UNC source path
    $vsLogicalFolder  = $vsLogicalFolder.
        ToLower() -replace "(^\s*[\\/]*)|([\\/]\s*$)", ""
    
    #####################################
    # Create file system folder path 
    # if it does not exist
    #
    $local:fileSystemPath = $vsTargetProjFilePath.
        SubString(0, $vsTargetProjFilePath.LastIndexOf("\")+1) +
        $vsLogicalFolder
    
    New-Item -Path $fileSystemPath -ItemType directory `
        -ErrorAction Ignore | Out-Null


    #####################################
    # If folder is in VS then exit
    #
    $local:xml = new-object System.Xml.XmlDocument
    $xml.Load($vsTargetProjFilePath)
    
    $local:vsItemsWithMatchingFolders = 
        ($xml.Project.ItemGroup.ChildNodes | 
        Where-Object { 
             $local:includePath = $_.GetAttribute("Include").ToLower()
             
             $includePath.Length -gt 0                                         `
             -and $FILE_EXT_TO_VS_NODE_MAP.Values -contains $_.Get_LocalName() `
             -and $includePath.StartsWith( $vsLogicalFolder )        `
        })

    $local:hasFilesInFolderWithinVs = (
        $vsItemsWithMatchingFolders | 
        Where-Object {
            $_.Include.Length -gt $vsLogicalFolder.length
        } | 
        Select -ExpandProperty Include
        ).length -gt 0
    

    if ($hasFilesInFolderWithinVs                 `
    -or ($vsItemsWithMatchingFolders.length -gt 0 `
         -and (! $ADD_EMPTY_FILE_TO_EMPTY_FOLDER      `
               -or [string]::IsNullOrWhiteSpace($EMPTY_FILENAME_STR)) ) ) {

            return
    }
    
    #####################################
    # Add item to filesystem
    #    
    if ($ADD_EMPTY_FILE_TO_EMPTY_FOLDER -and `
    ! [string]::IsNullOrWhiteSpace($EMPTY_FILENAME_STR)) {

        $local:txtFile = ($fileSystemPath + "\$EMPTY_FILENAME_STR")
        
        New-Item $txtFile -ItemType file -ErrorAction Ignore | Out-Null
    }
    
    #####################################
    # Add item to VS project
    #
    $vsLogicalFolder  # write-host
        
    $local:itemGroup = $xml.CreateElement("ItemGroup", $xml.Project.xmlns)    
    $xml.Project.AppendChild($itemGroup) | Out-Null

    $local:node = $null
    
    if ($ADD_EMPTY_FILE_TO_EMPTY_FOLDER `
    -and ![string]::IsNullOrWhiteSpace($EMPTY_FILENAME_STR)) {
    
        $node = $xml.CreateElement( 
            $FILE_EXT_TO_VS_NODE_MAP.Get_Item("default"), $xml.Project.xmlns)
            
        $node.setAttribute("Include", "$vsLogicalFolder\$EMPTY_FILENAME_STR" )
    }
    else {
        $node = $xml.CreateElement("Folder", $xml.Project.xmlns)
        $node.setAttribute("Include", $vsLogicalFolder )
    }
    
    $itemGroup.AppendChild($node) | Out-Null

    $xml.Save($vsTargetProjFilePath)

} # Add-EmptyFolderIntoProject()



################################################################################
# Update-FileContent
#
# Purpose: Update file content 
#
# Param:  $file - full path to file to update content
#  
#         $matchPattern - regular expression pattern for replacing content
# 
#         $replacePattern - match group pattern for replacing content
#
# Result: If file exists then it is updated with regular expression 
#         replace content; otherwise, no action is taken
#
function Update-FileContent ($file, $matchPattern, $replacePattern) {

    if (Test-Path -PathType Leaf -Path $file) {
    
        (Get-Content $file) -replace $matchPattern, $replacePattern `
            | Set-Content -Path $file | Out-Null
    }
} # Update-FileContent



################################################################################
# Update-CS-files
#
# Purpose: Update the following C-Sharp files
#
#             /App_Start/RouteConfig.cs - update namespace 
#                                         and remove generic MVC route
#
#             /global.asax              - update namespace
#
#             /global.asax.cs           - update namespace and
#                                         change inheritance from HttpApplication
#                                         Sitecore.Web.Application 
#
Function Update-CS-files($vsProjFile) {
    
    $local:xml = new-object System.Xml.XmlDocument
    $xml.Load($vsProjFile)
    
    $local:namespace = (Get-XmlNode -XmlDocument $xml `
       -NodePath "/Project/PropertyGroup/RootNamespace" `
       -NodeSeparatorCharacter "/" ).InnerText

    $local:rootProjFolder = $vsProjFile.SubString(0, 
        $vsProjFile.LastIndexOf("\")+1)
        
    $local:replacePattern = "`$1$namespace`$3"
    $local:namespacePattern = "(^\s*namespace\s+)([\w\d_]+)(\s*$)"

    #####################################
    # Update App_Start/RouteConfig.cs 
    # Namespace to match assembly and
    # generic route
    Update-FileContent -file "$rootProjFolder/App_Start/RouteConfig.cs" `
        -matchPattern $namespacePattern              `
        -replacePattern $replacePattern

    Update-FileContent -file "$rootProjFolder/App_Start/RouteConfig.cs" `
        -matchPattern "(.*url:\s*`")({controller\}/\{action\}/\{id\})(`".*)" `
        -replacePattern '$1$2/NO_WHERE$3'

    #####################################
    # Update Global.asax Namespace
    #
    Update-FileContent -file "$rootProjFolder/global.asax"                             `
                       -matchPattern "(^.*Inherits=`")([\w\d_]+)(\..*$)" `
                       -replacePattern $replacePattern

    #####################################
    # Update Global.asax.cs namespace and
    # parent class to:
    #    Sitecore.Web.Application
    #
    Update-FileContent -file "$rootProjFolder\global.asax.cs"  `
        -matchPattern $namespacePattern  `
        -replacePattern $replacePattern

    Update-FileContent -file "$rootProjFolder\global.asax.cs"                           `
        -matchPattern "(class\s+Global\s*:\s*)(HttpApplication)(\s*$)" `
        -replacePattern "`$1Sitecore.Web.Application`$3"
                       
} # Update-CS-files


################################################################################
# Get-PsScriptFolder()
# 
# Purpose: Get the full path to current running PowerScript file.
#           $PSScriptRoot is not available prior to PowerShell 3
#
# Param:   N/A
# 
# Return:  Full path to current PowserShell path
#
function Get-PsScriptFolder {

    # $PSScriptRoot is not available everywhere
    # return split-path -parent $MyInvocation.MyCommand.Definition
    return Split-Path -Parent $PSCommandPath
    
} # Get-PsScriptFolder()


################################################################################
# Find-VisualStudioTemplateProjectWithInPsFolder()
#
# Purpose:  Find visual studio template folder within the PowerShell
#            script path.  As a result, the user does not have to specify 
#           template path when calling the New-VsScProj command.
#
# Param:    $index - Because the script path can contain multiple 
#               VS project folders, the $index is used to select 
#                specific item.  
#
# Return:   If no project folder within script folder is found
#           then null is returned.
#
#           If the $index is not numeric or less than zero, then 
#           the first template folder is returned.
#
#            If the $index exceeds maximum number of template 
#           projects, then the last project path is returned.
#
#           Otherwise, the specified $indexed project path is
#           returned.
#
function Find-VisualStudioTemplateProjectWithInPsFolder($index) {

    $local:scriptPath = Get-PsScriptFolder
    
    $startFolder = $scriptPath 
    
    $local:vsProjFiles = (Get-ChildItem -Directory -Path $startFolder | 
        Where-Object {
            $local:subDirectory = $_
            Get-ChildItem -File -Path $subDirectory.FullName |
                Where-Object {
                    $_.Extension.ToLower() -eq ".csproj" `
                    -and $subDirectory.Name -imatch $VS_TEMPLATE_REGEX 
                }
        })
    
    if ($vsProjFiles -eq $null) {
        return $null
    }
    
    if ($index -isnot [int]) {
        $index = 0
    } else {
        $index = (0,$index | Measure -Maximum).Maximum
    }
    
    if ($vsProjFiles -is [System.Array]) {
        $index = (($vsProjFiles.length-1), $index | Measure -Minimum).Minimum
        return $vsProjFiles[$index].FullName
    }
    else {
        return $vsProjFiles.FullName
    }
} # Find-VisualStudioTemplateProjectWithInPsFolder()


################################################################################
# Expand-ZIPFile()
#
# Purpose:    Extract a zip file into a destination folder.
#             This provides a conveniant way to create a visual studio 
#            project based on a Sitecore distribution zip file.
#
# Param:  $zipSourceFile - full path to zip file to extract
#
#          $destination - destination folder to extract the zip file into
#
# Result: Extracted achive
#
# Reference: 
#    http://www.howtogeek.com/tips/how-to-extract-zip-files-using-powershell
#
function Expand-ZIPFile($zipSourceFile, $destination) {

    $local:shell = new-object -com shell.application
    $local:zip = $shell.NameSpace($zipSourceFile)
    
    New-Item -Path $destination -ItemType directory `
        -ErrorAction Ignore | Out-Null
    
    foreach($item in $zip.items()) {
        $shell.Namespace($destination).copyhere($item)
    }    
    
} # Expand-ZIPFile


################################################################################
# Get-XmlNode()
#
# Reference: 
#    http://blog.danskingdom.com/tag/selectsinglenode/
#    http://stackoverflow.com/questions/1766254/selectsinglenode-always-returns-null
#
function Get-XmlNode([System.Xml.XmlDocument]$XmlDocument, 
                     [string]$NodePath, 
                     [string]$NamespaceURI = "", 
                     [string]$NodeSeparatorCharacter = '.') {

    $local:nodes = Get-XmlNodes -XmlDocument $XmlDocument `
          -NodePath $NodePath         `
          -NamespaceUri $NamespaceURI `
          -NodeSeparatorCharacter $NodeSeparatorCharacter
    
    if ($nodes -is [System.Array]) {
        return $nodes[0]
    }
    return $nodes
} #Get-XmlNode

function Get-XmlNodes([System.Xml.XmlDocument]$XmlDocument, 
                     [string]$NodePath, 
                     [string]$NamespaceURI = "", 
                     [string]$NodeSeparatorCharacter = '.') {

    # If a Namespace URI was not given, 
    # use the Xml document's default namespace.
    if ([string]::IsNullOrEmpty($NamespaceURI)) { 
        $NamespaceURI = $XmlDocument.DocumentElement.NamespaceURI 
        }   
     
    # In order for SelectSingleNode() to actually work, 
    # we need to use the fully qualified node path along with 
    # an Xml Namespace Manager, so set them up.
    $local:xmlNsManager = New-Object `
        System.Xml.XmlNamespaceManager($XmlDocument.NameTable)
        
    $xmlNsManager.AddNamespace("ns", $NamespaceURI)
    
    $local:fullyQualifiedNodePath = 
        "$($NodePath.Replace($($NodeSeparatorCharacter), '/ns:'))"
     
    # Try and get the node, then return it. 
    # Returns $null if the node was not found.
    $local:node = $XmlDocument.
        SelectNodes($fullyQualifiedNodePath, $xmlNsManager)
        
    return $node
    
} # Get-XmlNodes


################################################################################
# Add-SolutionFiles()
#
# Purpose:  Add VS studio solution file and .nuget folder to the 
#           root of the project
#
# Params:   $vsTemplateFolder - source template folder to visual studio project.
#               The template folder and its parent folder will be searched
#               for a Visual Studio solution file (*.sln)
#
#            $vsTargetProjFilePath - Target visual studio project, to be added
#               in to the solution file
#
# Result:   The VS solution file is copied from the source template folder
#           or its parent folder to target project parent directory.
#           If the solution already exists, then no futher action is taken.
#           The solution file is renamed to match the project, and it is
#           updated to hold the project. 
#           Additionally, .nuget folder is also copied to the same location
#           as the solution file (e.g. parent folder of the project)
#
# Return: N/A
function Add-SolutionFiles($vsTemplateFolder, $vsTargetProjFilePath) {

    $local:parentTemplateFolder = $vsTemplateFolder.SubString(0,
        $vsTemplateFolder.Replace("/","\").LastIndexOf("\"))
    
    if (! $vsTargetProjFilePath.EndsWith(".csproj") -or
    ! (Test-Path -PathType Leaf -Path $vsTargetProjFilePath)) { 
        Throw [System.IO.FileNotFoundException] `
        "No visual studio project file detected at path $vsTargetProjFilePath"
    }

    $vsTargetProjFilePath = $vsTargetProjFilePath.Replace("/","\\")
    
    $local:lastSlashIndex = $vsTargetProjFilePath.LastIndexOf("\")
    
    $local:lastSlashIndex2= $vsTargetProjFilePath.LastIndexOf("\", `
        $lastSlashIndex-1, $lastSlashIndex)
        
    $local:parentTargetFolder=
        $vsTargetProjFilePath.SubString(0,$lastSlashIndex2)
    
    #####################################
    # Copy .nuget folder 
    #
    $local:srcNuget = "$parentTemplateFolder\.nuget"
    $local:dstNuget = "$parentTargetFolder\.nuget"
    
    if ( (Test-Path -PathType Container -Path $srcNuget ) `
    -and !(Test-Path -PathType Container -Path $dstNuget)) {

        Copy-Item -Recurse $srcNuget $dstNuget
    }
    
    #####################################
    # Copy solution file 
    #
    $local:srcSolutionFile = @($vsTemplateFolder, $parentTemplateFolder) |
        ForEach {
            Get-ChildItem -File $_ |
            Where-Object {$_.Extension.ToLower() -eq ".sln"} |
            Select -First 1 -ExpandProperty FullName
        }
    
    if ($srcSolutionFile -eq $null) {
        return
    }
    
    $local:projectName= (Get-ChildItem $vsTargetProjFilePath).BaseName
    
    $local:dstSolutionFile = "$parentTargetFolder\$projectName.sln"
    
    if (! (Test-Path -PathType Leaf -Path $dstSolutionFile)){
        Copy-Item $srcSolutionFile $dstSolutionFile
    }
    
    #####################################
    # Update Solution file
    Update-FileContent -file $dstSolutionFile `
        -matchPattern "(Project.*FAE04EC0-301F-11D3-BF4B-00C04F79EFBC.*)([\w\d\-_ ]*Template[\w\d\-_ ]*)(.*)\2(.*)\2(\.csproj.*)" `
        -replacePattern "`$1$projectName`$3$projectName`$4$projectName`$5" `

} # Add-SolutionFiles


################################################################################
# Update-WebConfigFile()
# 
# Purpose: Update web.config file assembly version and bindings to match
#          to that of the files within Sitecore bin folder
#
function Update-WebConfigFile($dstWebConfig) {
    
    $local:xml = new-object System.Xml.XmlDocument
    $xml.Load($dstWebConfig)

    $local:binFolder = $dstWebConfig.SubString(0, 
        $dstWebConfig.LastIndexOf("\")) + "\bin\"
        
    #####################################
    # Update the web.config assemblies 
    #
    Get-XmlNodes -XmlDocument $xml `
        -NodePath "/configuration/system.web/compilation/assemblies/add[@assembly]" `
        -NodeSeparatorCharacter "/" `
        | Foreach {
        
            $local:dllFile = $binFolder +
                $_.GetAttribute("assembly").split(",")[0] + ".dll"
                
            if (Test-Path -PathType Leaf -Path $dllFile) {
            
                $_.SetAttribute("assembly",
                    [Reflection.Assembly]::Loadfile($dllFile).
                    FullName)
            }
        }
        
    #####################################
    # Update web.config runtime bindings
    #
    $local:runtimeNode = Get-XmlNode  -XmlDocument $xml `
        -NodePath "/configuration/runtime" `
        -NodeSeparatorCharacter "." 
    
    $local:nsMgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsMgr.AddNamespace("ns","urn:schemas-microsoft-com:asm.v1")

    $runtimeNode.SelectNodes( "." +
                          "/ns:assemblyBinding" + 
                          "/ns:dependentAssembly" +
                          "/ns:assemblyIdentity[@publicKeyToken]" ,
                          $nsMgr `
        ) | Foreach {
        
            $local:assemblyIdentityNode = $_

            $local:bindingRedirectNode = 
                $_.ParentNode.SelectSingleNode("./ns:bindingRedirect", $nsMgr)
            
            $local:dllFile = $binFolder +
                $assemblyIdentityNode.GetAttribute("name").
                split(",")[0] + ".dll"
                
            if (Test-Path -PathType Leaf -Path $dllFile) {
                
                $local:assembly = 
                    [Reflection.Assembly]::Loadfile($dllFile).GetName()
                
                $local:matches=$null
                $assembly.fullname -imatch `
                    "PublicKeyToken=(?<publicKeyToken>[\w\d]+)" | Out-Null
                
                if (![string]::IsNullOrWhiteSpace($matches["PublicKeyToken"])) {
                    $assemblyIdentityNode.SetAttribute(
                        "publicKeyToken", $matches["PublicKeyToken"])
                }
                
                if ($bindingRedirectNode -ne $null) {
                
                    $local:newVersion = $assembly.Version.ToString()
                    
                    $bindingRedirectNode.SetAttribute("newVersion", $newVersion)
                        
                        
                    $local:oldVersion = $bindingRedirectNode.
                        GetAttribute("oldVersion")
                    if (![string]::IsNullOrWhiteSpace($oldVersion)) {
                        $bindingRedirectNode.SetAttribute("oldVersion",
                            ($oldVersion -replace ("-.*") ) + "-" + $newVersion)
                    }
                }
            }                   
        }
    
    
    $xml.save($dstWebConfig)
}

################################################################################
# Update-SitecoreVariable()
#
# Purpose: Update Sitecore web.config variables for node <sc.variable ...
#          This method is used to change the "dataFolder", as appose to 
#           changing /App_Config/Include/DataFolder.config file
#
# Param:   $webConfig - Full path to webConfig file
#
#          $varName - The var name matches the "name" attribute of SC variable
#                  <sc.variable name="varName" value=....
# 
#          $varValue - The value to set for "value" attribute of SC variable
#
# Return:  N/A
#
function Update-SitecoreVariable($webConfig, $varName, $varValue) {

        $local:xml = new-object System.Xml.XmlDocument
        $xml.load($webConfig)
        
        $local:sitecoreConfigs = (
            Get-XmlNodes -XmlDocument $xml `
            -NodePath "configuration/sitecore/sc.variable" `
            -NodeSeparatorCharacter "/")
        
        $local:node = $sitecoreConfigs | 
            Where-Object { 
                $_ -ne $null                                `
                -and $_.Get_LocalName() -eq "sc.variable"   `
                -and $_.getAttribute("name") -eq $varName   `
            }
        
        if ($node -ne $null) {

            $node.value = $varValue
            $xml.Save($webConfig)
        }
        
} # Update-SitecoreVariable()


#####################################
# Add-ReferenceToVsProject
#
# Purpose: Add a DLL library reference into VS project 
#
# Param: $vsProjFile - full path to visual studio project file
#
#        $referencePath - reference to DLL to include into VS
#
# Result: VS project file is update to have a reference to the
#         path specified only-if the referencePath file name is 
#         not already reference.
#
Function Add-ReferenceToVsProject ($vsProjFile, $referencePath) {

        $local:xml = new-object System.Xml.XmlDocument
        $xml.load($vsProjFile)
        
        $local:fileName = $referencePath.SubString(
            $referencePath.replace("/", "\").LastIndexOf("\") +1)
            
        $local:hasReference = (
            Get-XmlNode -XmlDocument $xml            `
            -NodePath "/Project/ItemGroup/Reference/HintPath" `
            -NodeSeparatorCharacter "/") |
            Where-Object {$_.InnerText -match $fileName.replace(".", "\.")} 
            
        if ($hasReference) {
            return
        }
        
        $referenceNode = (
            Get-XmlNode -XmlDocument $xml            `
            -NodePath "/Project/ItemGroup/Reference" `
            -NodeSeparatorCharacter "/")
            
        $referenceNode = $referenceNode.ParentNode.InsertBefore(
            $xml.CreateElement("Reference", $xml.Project.xmlns),
            $referenceNode)
        
        $referenceNode.SetAttribute("Include", "Sitecore.Kernel")
        
        $referenceNode.InnerXML = @"
            `n    
            <HintPath>$referencePath</HintPath>
            <SpecificVersion>False</SpecificVersion>
            <Private>False</Private>
            `n
"@
        $xml.Save($vsProjFile)
}


################################################################################
################################################################################
# New-VsScProj()
#
# Purpose: Main entry point to the script.  
#          A new visual studio project is created (or exiting one updated)
#          such that files from $source are added to the visual studio project
#          
function New-VsScProj {
    Param(
        [Parameter(Mandatory=$true, HelpMessage=
        "Sitecore zip file or extracted folder")]
        [string] $source,

        [Parameter(Mandatory=$false, HelpMessage=
        "Visual studio folder that serves as a template for a new project")]
        [string] $vsTemplateFolder,
        
        [Parameter(Mandatory=$false, HelpMessage=
        "Target path to new visual studio project folder. " +
        "If target project exists, then it is updated " +
        "(as apose to using the template).  If no " +
        "path is specified, then the current working directory is used")]
        [string] $vsTargetFolder,
        
        [Parameter(Mandatory=$false, HelpMessage=
        "Overwrite existing files")]
        [string] $overwriteExistingFiles = $false
    )

    # clean-up parameter values
    #
    $source = $source -replace "[`"']"    
    $vsTemplateFolder = $vsTemplateFolder -replace "[`"']"
    $vsTargetFolder = $vsTargetFolder -replace "[`"']"

    $local:scriptPath = Get-PsScriptFolder
    # Look for VS project template within the current folder if non-specified
    if ([string]::IsNullOrWhiteSpace($vsTargetFolder) `
        -or (Get-ChildItem -Path $vsTargetFolder -ErrorAction Ignore | 
             Where-Object {$_.Extension -ieq ".csproj"}) -eq $null
    ) {
        $vsTemplateFolder = Find-VisualStudioTemplateProjectWithInPsFolder
    }
        
    $local:templateFolderName = $vsTemplateFolder.
        SubString($vsTemplateFolder.LastIndexOf("\")+1)
    
    #####################################
    # Set target folder to current 
    # folder if non specified.  Also,
    # add a 
    #
    if ([string]::IsNullOrWhiteSpace($vsTargetFolder)) {
    
        $local:i = 0
        do {
            ++$i
            
            $vsTargetFolder = (Convert-Path .) +
                ("\$templateFolderName$i" -replace $VS_TEMPLATE_REGEX,"")
                
        } while( Test-Path -PathType Container -Path $vsTargetFolder)    
    }

    
    #####################################
    # if sourceFolder is a Zip file then 
    # we'll 1st extrat it into temp dir 
    #
    $local:scRootFolder = $source
    
    $local:targetZipFolder = $null
    
    if ((Test-Path -PathType Leaf -Path $source) `
    -and $source.toLower().EndsWith(".zip")) {

        do {
            $targetZipFolder = [System.IO.Path]::GetTempPath() +
                 [guid]::NewGuid().ToString().SubString(0,8)
        } while (Test-Path -PathType Leaf -Path $targetZipFolder)


        Write-Host "Extracting zipfile into $targetZipFolder"
        
        Expand-ZIPFile $source $targetZipFolder
        
        $scRootFolder = Get-ChildItem -Path $targetZipFolder -ErrorAction ignore `
            | Where-Object {$_.PSIsContainer}    `
            | Select-Object -Last 1 -ExpandProperty FullName

        if ([string]::IsNullOrWhiteSpace($scRootFolder)) {
            $scRootFolder = $targetZipFolder
        }
    }
    
    #####################################
    # drill down into subfolders until 
    # we can confirm that $scRootFolder 
    # contains folders "Data" and "Website"
    #
    
    if (!(Test-Path -PathType Container -Path $scRootFolder)) {
        throw [System.IO.FileNotFoundException]        `
            "Invalid Sitecore root folder $scRootFolder"    
            return
    }
    
    $local:subFolders = Get-ChildItem -Directory -Path $scRootFolder         
    
    $local:isScRootFolder =  ($subFolders | Where-Object {
        $_.Name -ieq "Data" -or $_.Name -ieq "Website" 
        }).count -gt 1
    
    while (!$isScRootFolder -and $subFolders.count -gt 0) {
    
        $scRootFolder = ($subFolders | Select -First 1 ).FullName
        
        $subFolders = Get-ChildItem -Directory -Path $scRootFolder         
        
        $isScRootFolder =  ($subFolders | Where-Object {
            $_.Name -contains "Data" -or $_.Name -contains "Website"
            }).count -gt 1
    } 


    #####################################
    # create the visual studio project
    #
    if ($isScRootFolder) {
    
        $local:hasTemplate = 
            ![string]::IsNullOrWhiteSpace( $vsTemplateFolder)
            
        if ($hasTemplate) {
            Copy-VsTemplateFolder       `
                $vsTemplateFolder $vsTargetFolder   | Out-Null
        }
        
        $local:vsProjFile = Get-VsProjectFilePath $vsTargetFolder
        
        Add-TypeScriptSupportToVs $vsProjFile
        
        Copy-FilesIntoProject                   `
            -vsTargetProjFilePath $vsProjFile   `
            -srcFolder "$scRootFolder\Website"  `
            -vsLogicalFolder "\"                `
            -overwriteExistingFiles $overwriteExistingFiles
            
        Copy-FilesIntoProject                   `
            -vsTargetProjFilePath $vsProjFile   `
            -srcFolder "$scRootFolder\Data"     `
            -vsLogicalFolder "\App_Data"        `
            -exclusionMatchFilter ".*\\indexes\\.*"  `
            -overwriteExistingFiles $overwriteExistingFiles

        if ($hasTemplate) {
            Write-Host "Add to Solution $vsProjFile"
            Add-SolutionFiles  $vsTemplateFolder $vsProjFile
        }
        else {
            Write-Host "Solution file is not updated"
        }

        
        #####################################
        # Overwrite template web.config
        #
        $local:srcWebConfig = "$scRootFolder\Website\web.config" 
        $local:dstWebConfig = "$vsTargetFolder\web.config"
        $local:templateWebConfig = "$vsTemplateFolder\web.config"
        
        $local:isTemplateWebConfig = $hasTemplate -and `
            ((Compare-Object -ReferenceObject (Get-Content $dstWebConfig) `
                 -DifferenceObject (Get-Content $templateWebConfig) `
            ) -eq $null)
            
        if ($isTemplateWebConfig) {
            Copy-Item $srcWebConfig $dstWebConfig -Force
        }
        
        Update-WebConfigFile $dstWebConfig

        Update-SitecoreVariable -webConfig $dstWebConfig `
            -varName "dataFolder" -varValue "/App_Data"
        
        
        Update-CS-files -vsProjFile $vsProjFile
        
        Add-ReferenceToVsProject -vsProjFile $vsProjFile `
            -referencePath "./bin/Sitecore.Kernel.dll"
        
        #####################################
        # Add empty folders to FileSystem 
        # and Visual studio
        #
        @("\temp", "\upload", "\App_Data\Indexes") |
            ForEach { 
                Add-EmptyFolderIntoProject -vsTargetProjFilePath $vsProjFile `
                       -vsLogicalFolder $_ 
            }
        
            
    } # if ($isScRootFolder) 
    
    
    #####################################
    # Cleanup 
    #
    if ($targetZipFolder -ne $null `
    -and (Test-Path -PathType Container $targetZipFolder) -eq $true ) {
    
        Write-Host "Deleting $targetZipFolder"
        
        Remove-Item -Recurse $targetZipFolder
    }
    
    if (!$isScRootFolder) {
    
        throw [System.IO.FileNotFoundException]        `
            "Sitecore root folder $scRootFolder does " +
            "not contain 'Website' and 'Data' folders"
    }
    
} # New-VsScProj()


################################################################################
################################################################################
# MAIN
#

if ( [string]::IsNullOrWhiteSpace($source)) {

$local:scriptName = Split-Path -Leaf $PSCommandPath
@"
   `nMissing command line arguments!!!
    
   $scriptName -sourceFolder <p1> -vsTemplateFolder <p2> -$vsTargetFolder <p3>

       -soruce (required):
            Sitecore root folder that contains Data, Website, and 
            Database folders.  The sourceFolder may also be 
            Sitecore distribution zip file. 

       -vsTemplateFolder (optional):
            visual studio project that is to be used as a 
            template for creating a new project that contains the 
            sitecore files. If none specified, then the powershell
            script folder is searched for a template folder.

       -vsTargetFolder (optional):
            target destination for new visual studio project.  
            As convention, the leaf folder name is used as the 
            project name, assembly and namespace.
            
       -overwriteExistingFiles (default true):
            If set to true, then existing files in the target folder 
            are overwriten during the copy process.  Otherwise,
            the will not be altered.
    
"@

    # prmopt the user for the required source
    #
    New-VsScProj -vsTemplateFolder $vsTemplateFolder `
        -vsTargetFolder $vsTargetFolder            `
        -overwriteExistingFiles $overwriteExistingFiles
}
else {
    New-VsScProj -source $source              `
        -vsTemplateFolder $vsTemplateFolder `
        -vsTargetFolder $vsTargetFolder     `
        -overwriteExistingFiles $overwriteExistingFiles

}

$source = 
$vsTemplateFolder =
$vsTargetFolder = 
$overwriteExistingFiles = 
$null

# Main