---
title:  'Convert OneNote to MarkDown'
author:
- Sjoerd de Valk, SPdeValk Consultancy
- modified by nixsee, a guy
date: 2020-07-14 22:35:00
keywords: [migration, tooling, onenote, markdown, powershell]
abstract: |
  This document is about converting your OneNote data to Markdown format.
permalink: /index.html
---

Credit for this script goes to the wizard @SjoerdV who created the original script [here](https://github.com/SjoerdV/ConvertOneNote2MarkDown). I've taken it and made a variety of modifications and improvements.

# Convert OneNote to MarkDown

## Summary

Ready to make the step to Markdown and saying farewell to your OneNote, EverNote or whatever proprietary note taking tool you are using? Nothing beats clear text, right? Read on!

The powershell script 'ConvertOneNote2MarkDown-v2.ps1' will utilize the OneNote Object Model on your workstation to convert all OneNote pages to Word documents and then utilizes PanDoc to convert the Word documents to Markdown (.md) format. It will also:

* Create a **folder structure** for your Notebooks and Sections
* Process pages that are in sections at the **Notebook, Section Group and 1st Nested Section Group levels**
* Allow you to **choose between creating subfolders for subpages** (e.g. Page\Subpage.md) or **appending prefixes** (e.g. Page_Subpage.md)
* Allow you you choose between putting all **Images** in a central '/media' folder for each notebook, or in a separate '/media' folder in each folder of the hierarchy
* Fix image references in the resulting .md files, generating *relative* references to the image files within the markdown document
* Extract all **File Objects** to the same folder as where the page is in and fix references in the resulting .md files
* Allow you to select between **discarding or keeping intermediate Word files**
* Allow user can **select which markdown format will be used**, defaulting to Pandoc's standard format, which strips any HTML from tables along with other desirable (for me) formatting choices.
   * markdown (Pandoc’s Markdown)
   * commonmark (CommonMark Markdown)
   * gfm (GitHub-Flavored Markdown), or the deprecated and less accurate markdown_github; use markdown_github only if you need extensions not supported in gfm.
   * markdown_mmd (MultiMarkdown)
   * markdown_phpextra (PHP Markdown Extra)
   * markdown_strict (original unextended Markdown)
* See more details on these options here: https://pandoc.org/MANUAL.html#options
## Known Issues

1. If there are any collapsed paragraphs in your pages, the collapsed/hidden paragraphs will not be exported
    * You can use the included Onetastic Macro script to automatically expand all paragraphs in each Notebook 
    * [Download Onetastic here](https://getonetastic.com/download) and, once installed,  double click the macro file to install it within Onetastic
1. Password protected sections should be unlocked before continuing, the Object Model does not have access to them if you don't
1. You should start by 'flattening' all pen/hand written elements in your onennote pages. Because OneNote does not have this function you will have to take screenshots of your pages with pen/hand written notes and paste the resulting image and then remove the scriblings. If you are a heavy 'pen' user this is a very cumbersome. **If you have an automated solution for this, please let me know**
1. Relative paths can not be used as input for the target folder. Always use an absolute path (ex. 'c:\temp\notes').
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
1. Open a PowerShell terminal (as Administrator) and navigate to the folder containing the script and run it (:

    ```'.\ConvertOneNote2MarkDown-v2.ps1'```
    
    if you receive an error, try running this line to bypass security:
     "Set-ExecutionPolicy Bypass -Scope Process"
    
1. It will ask you for the path to store the markdown folder structure. Please use an empty folder.

    **Attention:** use a full absolute path for the destination
1. It will ask you whether you want to create subfolders (1) or append prefixes (2) for subpages.

1. It will ask you which conversion method/markdown format you want: 1-6, defaulting to 1: Pandoc

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

### [2.1] - 2020-07-15
#### Added
* Prompt for keep or discard .docx files
* Prompt to have images in central folder or separate ones for each folder in hierarchy

#### Changed
* User prompt layouts 

# Removed
* Nothing


### [2.0] - 2020-07-14
#### Added
* Consolidated prior scripts into one
* Prompt for markdown format selection
* Prompt to choose between prefix and subfolders for subpages

#### Changed
* Now produces relative references to images (e.g ../../media
* Each notebook has a centralized images/media folder

#### Removed
* Extraneous code

### [1.1] - 2020-07-11
#### Added
 * two new scripts to allow for Section Groups, as well as Section Groups + Subfolders for Subpages
 
#### Changed
* Pandoc instead of gfm set as default format

#### Removed
* Nothing

### [1.0.0] - 2019-05-19

#### Added

* Initial Release

#### Changed

* Nothing

#### Removed

* Nothing


## Credits 
* Avi Aryan for the awesome [VSCodeNotebook](https://github.com/aviaryan/VSCodeNotebook) port
* @SjoerdV for the original script
