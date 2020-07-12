---
title:  'Convert OneNote to MarkDown'
author:
- Sjoerd de Valk, SPdeValk Consultancy
- modified by nixsee, a guy
date: 2019-05-19 22:35:00
last_modified_at: 2020-11T00:41:58+02:00
keywords: [migration, tooling, onenote, markdown, powershell]
abstract: |
  This document is about converting your OneNote data to Markdown format.
permalink: /index.html
---

Full credit for this script goes to the wizard @SjoerdV who created the original script [here](https://github.com/SjoerdV/ConvertOneNote2MarkDown). I've simply made a couple variations of his script that allow for section groups and creation of subfolders rather than prefixes for subpages.

# Convert OneNote to MarkDown

[![Github All Releases](https://img.shields.io/github/downloads/SjoerdV/ConvertOneNote2MarkDown/total.svg)](https://github.com/SjoerdV/ConvertOneNote2MarkDown/releases)

## Summary

!!! question Ready to make the step to Markdown and saying farewell to your OneNote, EverNote or whatever proprietary note taking tool you are using? Nothing beats clear text, right? Read on!

The powershell script 'ConvertOneNote2MarkDown.ps1' will utilize the OneNote Object Model on your workstation to convert all OneNote pages to Word documents and then utilizes PanDoc to convert the Word documents to Markdown (.md) format. It will also:

* Create a **folder structure** for your Notebooks and Sections
* Append **prefixes to page filenames** if they were indented beneath other pages (so called 'page levels')
  * script **"ConvertOneNote2MarkDownSectionGroupsSubpageFolders.ps1"** will **create subfolders** rather than add prefixes
* Extract all **Images** to the '/media' folder of each section and fix references in the resulting .md files
  * this can be annoying when using the above-mentioned subfolder script, but I don't know enough to create a single media folder
* Extract all **File Objects** to the same folder as where the page is in and fix references in the resulting .md files
* Cleanup intermediate Word files
* Script **ConvertOneNote2MarkDownSectionGroups.ps1** will allow you to work with section groups that are at the root/top-level of the notebook. It does not extract any sections that are not in section groups, so just make a dummy group for any top-level sections.

## Known Issues

1. Password protected sections should be unlocked before continuing, the Object Model does not have access to them if you don't
1. ~~Section Groups on the first level are listed but are ignored. Nested Section Groups are not processed at all.~~
    * ~~Recommendation: if you make heavy use of (Nested) Section Groups you first have to reorganize in a way that they are out of the picture. Usually creating a new Notebook named the same as your Section Group and moving all relevant Sections.~~
    
1. You should start by 'flattening' all pen/hand written elements in your onennote pages. Because OneNote does not have this function you will have to take screenshots of your pages with pen/hand written notes and paste the resulting image and then remove the scriblings. If you are a heavy 'pen' user this is a very cumbersome. **If you have an automated solution for this, please let me know**
1. Relative paths can not be used as input for the target folder. Always use an absolute path (ex. 'c:\temp\notes').
1. This script uses only absolute paths internally, mainly because pandoc on Windows has trouble processing relative paths and for consistency. This will not be changed.
1. While running the conversion OneNote will be unusable and it is recommended to 'walk away' and have some coffee as the Object Model might be interrupted if you do anything else.
1. Linked file object in .md files are clickable in VSCode, but do not open in their associated program, you will have to open the files directly from the file system.
1. Anything I did not catch... please submit an issue.

## Requirements

* Windows >= 10

  * I have only tested this on Windows...

* Microsoft OneNote >= 2016

* Microsoft Word >= 2016

* PanDoc >= 2.7.2

  * TIP: Use [Chocolatey](https://chocolatey.org/docs/installation#install-with-powershellexe) to install Pandoc on Windows, this will also set the right path (environment) statements. (https://chocolatey.org/packages/pandoc)
    

## Installation

Clone this repository to acquire the powershell script.

## Usage

1. Start the OneNote application as Administrator. All notebooks currently loaded in OneNote will be converted
1. Open a PowerShell terminal (as Administrator) and navigate to the folder containing the script and run it:

    ```powershell
    '.\ConvertOneNote2MarkDown.ps1'
    "```"
    
* if you receive an error, try running this line to bypass security:
     "Set-ExecutionPolicy Bypass -Scope Process"

1. It will ask you for the path to store the markdown folder structure. Please use an empty folder.

    **Attention:** use a full absolute path for the destination

1. Sit back and wait until the process completes

## Results

The script will log any errors encountered at the end of its run, so please review, fix and run again if needed.
If you are satisfied check the results with a markdown editor like VSCode. All images should popup just right in the Preview Pane for Markdown files.

## Recommendations

1. I would like to recommend this repository [VSCodeNotebook](https://github.com/aviaryan/VSCodeNotebook) to host your resulting Markdown Notes folder structure. This solution supports encrypting sensitive (markdown) files and works quite nicely.
1. While working with markdown in VSCode these are the extensions I like using:

```powershell
    .\code `
    --install-extension davidanson.vscode-markdownlint `
    --install-extension ms-vscode.powershell-preview `
    --install-extension jebbs.markdown-extended `
    --install-extension telesoho.vscode-markdown-paste-image `
    --install-extension redhat.vscode-yaml `
    --install-extension vscode-icons-team.vscode-icons `
    --install-extension ms-vsts.team
```

> NOTE: The bottom three are not really markdown related but are quite obvious.

## Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

### [Unreleased]

### [1.0.0] - 2019-05-19

#### Added

* Initial Release

#### Changed

* Nothing

#### Removed

* Nothing

## Credits

* Avi Aryan for the awesome [VSCodeNotebook](https://github.com/aviaryan/VSCodeNotebook) port
