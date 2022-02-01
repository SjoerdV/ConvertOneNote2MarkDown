
<#
.SYNOPSIS

Convert OneNote to Markdown Tool
Copyright (C) 2021  Sjoerd de Valk

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3 as published by
the Free Software Foundation.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
.DESCRIPTION

This script will convert your OneNote notebook, sections and pages to Markdown files using the OneNote InterOp API.
Reference: https://matthewsetter.com/technical-documentation/asciidoc/convert-markdown-to-asciidoc-with-kramdoc/
Github: https://github.com/asciidoctor/kramdown-asciidoc
.EXAMPLE

Run this script file as follows: .\ConvertOneNote2MarkDown.ps1
#>


Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true)]
    [String]$Name
  )
  $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
  return ($newName.Replace(" ", "_"))
}

# ask for the Notes root path
$notesdestpath = Read-Host -Prompt "Enter the (preferably empty!) folder path (without trailing backslash!) that will contain your resulting Notes structure. ex. 'c:\temp\notes'"

if (Test-Path -Path $notesdestpath) {
  # open OneNote hierarchy
  $OneNote = New-Object -ComObject OneNote.Application
  [xml]$Hierarchy = ""
  $totalerr = ""
  $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

  foreach ($notebook in $Hierarchy.Notebooks.Notebook) {
    #if ($notebook.Name -eq "Company X Notebook" ) {
      " "
      $notebook.Name
      $notebookFileName = "$($notebook.Name)" | Remove-InvalidFileNameChars
      New-Item -Path "$($notesdestpath)\" -Name "$($notebookFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
      "=============="

      foreach ($sectiongroup in $notebook.SectionGroup) {
        if ($sectiongroup.isRecycleBin -ne 'true') {
          "## " + $sectiongroup.Name
        }
      }
      foreach ($section in $notebook.Section) {
        #if ($section.Name -eq "Adjust Forms and Workflow Section") {
          "--------------"
          "### " + $section.Name
          $sectionFileName = "$($section.Name)" | Remove-InvalidFileNameChars
          New-Item -Path "$($notesdestpath)\$($notebookFileName)" -Name "$($sectionFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
          [int]$previouspagelevel = 1
          [string]$previouspagenamelevel1 = ""
          [string]$previouspagenamelevel2 = ""
          [string]$pageprefix = ""
          $pageRecurrence = 1
          foreach ($page in $section.Page) {
            "#### " + $page.name
            #if ($page.name -eq "Workflow Documentation Page") {
              # set page variables
              $pagelevel = $page.pagelevel
              $pagelevel = $pagelevel -as [int]
              $pageid = ""
              $pageid = $page.ID
              $pagename = ""
              $pagename = $page.name | Remove-InvalidFileNameChars
              $fullexportdirpath = ""
              $fullexportdirpath = "$($notesdestpath)\$($notebookFileName)\$($sectionFileName)"
              $fullexportpathwithoutextension = ""
              $fullexportpath = ""

              # determine right name prefix based on pagelevel
              if ($pagelevel -eq 1) {
                $pageprefix = ""
                $previouspagenamelevel1 = $pagename
                $previouspagenamelevel2 = ""
                $previouspagelevel = 1
              }
              elseif ($pagelevel -gt $previouspagelevel) {
                if ($pagelevel -eq 2) {
                  $pageprefix = "$($previouspagenamelevel1)"
                  $previouspagenamelevel2 = $pagename
                  $previouspagelevel = 2
                }
                if ($pagelevel -eq 3) {
                  $pageprefix = "$($previouspagenamelevel1)_$($previouspagenamelevel2)"
                  $previouspagelevel = 3
                }
              }
              elseif ($pagelevel -eq $previouspagelevel -and $pagelevel -ne 1) {
                if ($pagelevel -eq 2) {
                  $pageprefix = "$($previouspagenamelevel1)"
                  $previouspagenamelevel2 = $pagename
                }
                if ($pagelevel -eq 3) {
                  $pageprefix = "$($previouspagenamelevel1)_$($previouspagenamelevel2)"
                }
              }
              elseif ($pagelevel -lt $previouspagelevel -and $pagelevel -ne 1) {
                if ($pagelevel -eq 2) {
                  $pageprefix = "$($previouspagenamelevel1)"
                  $previouspagenamelevel2 = $pagename
                  $previouspagelevel = 2
                }
              }
              if ($pageprefix) {
                $pagename = "$($pageprefix)_$($pagename)"
              }

              $fullexportpathwithoutextension = "$($fullexportdirpath)\$($pagename)"
              $fullexportpath = "$($fullexportpathwithoutextension).docx"

              # in case multiple pages with the same name exist in a section, postfix the filename
              if ([System.IO.File]::Exists("$($fullexportpathwithoutextension).md")) {
                $pagename = "$($pagename)_$pageRecurrence"
                $pageRecurrence++
              }

              $fullexportpathwithoutextension = "$($fullexportdirpath)\$($pagename)"
              $fullexportpath = "$($fullexportpathwithoutextension).docx"

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
                pandoc.exe -f docx -t gfm -i $fullexportpath -o "$($fullexportpathwithoutextension).md" --wrap=none --markdown-headings=atx --extract-media="$($fullexportdirpath)"
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
                  ((Get-Content -LiteralPath "$($fullexportpathwithoutextension).md" -Raw).Replace("$($pageinsertedfile.InsertedFile.preferredName)", "[$($destfilename)](./$($destfilename))")) | Set-Content -LiteralPath "$($fullexportpathwithoutextension).md"
                }
                catch {
                  Write-Host "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                  $totalerr += "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
              }

              # rename images
              $timeStamp = (Get-Date -Format o).ToString()
              $timeStamp = $timeStamp.replace(':', '')
              $re = [regex]"\d{4}-\d{2}-\d{2}T"
              $images = Get-ChildItem -Path "$($fullexportdirpath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch $re }
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
                # Change MD file Image Name References
                try {
                  ((Get-Content -LiteralPath "$($fullexportpathwithoutextension).md" -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -LiteralPath "$($fullexportpathwithoutextension).md"
                }
                catch {
                  Write-Host "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                  $totalerr += "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
              }

              # change MD file Image Path References
              try {
                # Change MD file Image Path References in Markdown
                ((Get-Content -LiteralPath "$($fullexportpathwithoutextension).md" -Raw).Replace("$($fullexportdirpath.Replace("\","\\"))/", "")) | Set-Content -LiteralPath "$($fullexportpathwithoutextension).md"
                # Change MD file Image Path References in HTML
                ((Get-Content -LiteralPath "$($fullexportpathwithoutextension).md" -Raw).Replace("$($fullexportdirpath)/", "")) | Set-Content -LiteralPath "$($fullexportpathwithoutextension).md"
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
            #}
          }
        #}
      }
    #}
  }
  # release OneNote hierarchy
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
  Remove-Variable OneNote
  $totalerr
}
else {
  Write-Host "This path is NOT valid"
}
