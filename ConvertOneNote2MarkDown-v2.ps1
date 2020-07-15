Function Remove-InvalidFileNameChars {
    param(
        [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [String]$Name
    )
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
    return (((($newName -replace "\s", "_") -replace "\[", "(") -replace "\]", ")").Substring(0,$(@{$true=130;$false=$newName.length}[$newName.length -gt 130])))
}
  
Function ProcessSections ($group, $FilePath) {
    foreach ($section in $group.Section) {
        "--------------"
        "### " + $section.Name
        $sectionFileName = "$($section.Name)" | Remove-InvalidFileNameChars
        New-Item -Path "$($FilePath)" -Name "$($sectionFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
        [int]$previouspagelevel = 1
        [string]$previouspagenamelevel1 = ""
        [string]$previouspagenamelevel2 = ""
        [string]$pageprefix = ""
        
        foreach ($page in $section.Page) {
            # set page variables
            $recurrence = 1
            $pagelevel = $page.pagelevel
            $pagelevel = $pagelevel -as [int]
            $pageid = ""
            $pageid = $page.ID
            $pagename = ""
            $pagename = $page.name | Remove-InvalidFileNameChars
            $fullexportdirpath = ""
            $fullexportdirpath = "$($FilePath)\$($sectionFileName)"
            $fullfilepathwithoutextension = ""
            $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
            $fullexportpath = ""
            $fullexportpath = "$($fullfilepathwithoutextension).docx"

            # make sure there is no existing Word file
            if ([System.IO.File]::Exists($fullexportpath)) {
                try {
                    Remove-Item -path $fullexportpath -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())`r`n"
                }
            }

            # in case multiple pages with the same name exist in a section, postfix the filename
            if ([System.IO.File]::Exists("$($fullfilepathwithoutextension).md")) {
                $pagename = "$($pagename)_$recurrence"
                $recurrence++
            }
            
            # process for subpage prefixes
            if ($pagelevel -eq 1) {
                $pageprefix = ""
                $previouspagenamelevel1 = $pagename
                $previouspagenamelevel2 = ""
                #$previouspagelevel = 1
                "#### " + $page.name
            }
            elseif ($pagelevel -eq 2) {
                    $pageprefix = "$($previouspagenamelevel1)"
                    $previouspagenamelevel2 = $pagename
                    $previouspagelevel = 2
                    "##### " + $page.name
            }
            elseif ($pagelevel -eq 3) {
                    $pageprefix = "$($previouspagenamelevel1)$($prefixjoiner)$($previouspagenamelevel2)"
                    $previouspagelevel = 3
                    "####### " + $page.name
            }
            
            #if level 2 or 2 (has pageprefix)
            if ($pageprefix) {
                #create filename prefixes and filepath if prefixes selected
                if ($prefixFolders -eq 2) {
                    $pagename = "$($pageprefix)_$($pagename)"
                    $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                }
                #all else/default, create subfolders and filepath if subfolders selected
                else {
                    New-Item -Path "$($fullexportdirpath)\$($pageprefix)" -ItemType "directory" -ErrorAction SilentlyContinue
                    $fullexportdirpath = "$($fullexportdirpath)\$($pageprefix)"
                    $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                    $levelsprefix = "../"*($levelsfromroot+$pagelevel-1)+".."
                }
            }
            else {
                $levelsprefix = "../"*($levelsfromroot)+".."
            }

            # publish OneNote page to Word
            try {
                $OneNote.Publish($pageid, $fullexportpath, "pfWord", "")
            }
            catch {
                Write-Host "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())`r`n"
            }

            # convert Word to Markdown
            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            try {
                pandoc.exe -f docx -t $converter -i $fullexportpath -o "$($fullfilepathwithoutextension).md" --wrap=none --atx-headers --extract-media="$($NotebookFilePath)"
            }
            catch {
                Write-Host "Error while converting file '$($page.name)' to md: $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error while converting file '$($page.name)' to md: $($Error[0].ToString())`r`n"
            }

            # export inserted file objects
            [xml]$pagexml = ""
            $OneNote.GetPageContent($pageid, [ref]$pagexml, 7)
            $pageinsertedfiles = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $_.InsertedFile }
            foreach ($pageinsertedfile in $pageinsertedfiles) {
                $destfilename = ""
                try {
                    $destfilename = $pageinsertedfile.InsertedFile.preferredName | Remove-InvalidFileNameChars
                    Copy-Item -Path "$($pageinsertedfile.InsertedFile.pathCache)" -Destination "$($fullexportdirpath)\$($destfilename)" -Force
                }
                catch {
                    Write-Host "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())`r`n"
                }
                # Change MD file Object Name References
                try {
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($pageinsertedfile.InsertedFile.preferredName)", "[$($destfilename)](./$($destfilename))")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                }
                catch {
                    Write-Host "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }

            # rename images to have unique timestamp names
            $timeStamp = (Get-Date -Format o).ToString()
            $timeStamp = $timeStamp.replace(':', '')
            $re = [regex]"\d{4}-\d{2}-\d{2}T"
            $images = Get-ChildItem -Path "$($NotebookFilePath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch $re }
            foreach ($image in $images) {
                $newimageName = "$($image.BaseName)_$($timeStamp)$($image.Extension)"
                # Rename Image
                try {
                    Rename-Item -Path "$($image.FullName)" -NewName $newimageName -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Error while renaming image '$($image.FullName)' for page '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming image '$($image.FullName)' for page '$($page.name)': $($Error[0].ToString())`r`n"
                }
                # Change MD file Image filename References
                try {
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                }
                catch {
                    Write-Host "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }

            # change MD file Image Path References
            try {
                # Change MD file Image Path References in Markdown
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($NotebookFilePath.Replace("\","\\"))", "$($levelsprefix)")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                # Change MD file Image Path References in HTML
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($NotebookFilePath)", "$($levelsprefix)")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
            }
            catch {
                Write-Host "Error while renaming image file path references for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error while renaming image file path references for file '$($page.name)': $($Error[0].ToString())`r`n"
            }

            # Cleanup Word files
            try {
                Remove-Item -path "$fullexportpath" -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())`r`n"
            }
        }
    }
}

# ask for the Notes root path
$notesdestpath = Read-Host -Prompt "Enter the (preferably empty!) folder path (without trailing backslash!) that will contain your resulting Notes structure. ex. 'c:\temp\notes'"

# prompt for prefix vs subfolders
[Int]$prefixFolders = Read-Host -Prompt "Press 1 to create subfolders for subpages (e.g. Page\Subpage.md ). Press 2 to create prefixes for subpages (Page_Subpage.md). Defaults to 1 on other input."
if ($prefixFolders -eq 2) {
    $prefixFolders = 2 
    $prefixjoiner = "_"
}
else {
    $prefixFolders = 1
    $prefixjoiner = "\"
}

#prompt for conversion type
"Select conversion type"
"-----------------------------------------------"
"1: markdown (Pandoc)"
"2: commonmark (CommonMark Markdown)"
"3: gfm (GitHub-Flavored Markdown)"
"4: markdown_mmd (MultiMarkdown)"
"5: markdown_phpextra (PHP Markdown Extra)"
"6: markdown_strict (original unextended Markdown)"
[int]$conversion = Read-Host -Prompt "Select 1-6 (Default 1 on other entry/error): "
if ($conversion -eq 2){ $converter = "commonmark"}
elseif ($conversion -eq 3){ $converter = "gfm"}
elseif ($conversion -eq 4){ $converter = "markdown_mmd"}
elseif ($conversion -eq 5){ $converter = "markdown_phpextra"}
elseif ($conversion -eq 6){ $converter = "markdown_strict"}
else { $converter = "markdown"}

if (Test-Path -Path $notesdestpath) {
    # open OneNote hierarchy
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$Hierarchy = ""
    $totalerr = ""
    $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

    foreach ($notebook in $Hierarchy.Notebooks.Notebook) {
        " "
        $notebook.Name
        $notebookFileName = "$($notebook.Name)" | Remove-InvalidFileNameChars
        New-Item -Path "$($notesdestpath)\" -Name "$($notebookFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
        $NotebookFilePath = "$($notesdestpath)\$($notebookFileName)"
        $levelsfromroot = 0
        "=============="
        #process any sections that are not in a section group
        ProcessSections $notebook $NotebookFilePath
        
        #start looping through any top-level section groups in the notebook
        foreach ($sectiongroup1 in $notebook.SectionGroup) {
            $levelsfromroot = 1
            if ($sectiongroup1.isRecycleBin -ne 'true') {
                "# " + $sectiongroup1.Name
                $sectiongroupFileName1 = "$($sectiongroup1.Name)" | Remove-InvalidFileNameChars
                New-Item -Path "$($notesdestpath)\$($notebookFileName)" -Name "$($sectiongroupFileName1)" -ItemType "directory" -ErrorAction SilentlyContinue 
                $sectiongroupFilePath1 =  "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)"
                ProcessSections $sectiongroup1 $sectiongroupFilePath1

                #start looping through any 2nd level section groups within the 1st level section group
                foreach ($sectiongroup2 in $sectiongroup1.SectionGroup) {
                    $levelsfromroot = 2
                    if ($sectiongroup2.isRecycleBin -ne 'true') {
                        "## " + $sectiongroup2.Name
                        $sectiongroupFileName2 = "$($sectiongroup2.Name)" | Remove-InvalidFileNameChars
                        New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)" -Name "$($sectiongroupFileName2)" -ItemType "directory" -ErrorAction SilentlyContinue
                        $sectiongroupFilePath2 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)"
                        ProcessSections $sectiongroup2 $sectiongroupFilePath2
                    }
                }
            }
        }        
    }
    
    # release OneNote hierarchy
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
    Remove-Variable OneNote
    $totalerr
}
else {
Write-Host "This path is NOT valid"
}
