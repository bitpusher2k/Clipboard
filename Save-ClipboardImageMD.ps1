#           Bitpusher
#            \`._,'/
#            (_- -_)
#              \o/
#          The Digital
#              Fox
#    https://theTechRelay.com
# https://github.com/bitpusher2k
#
# Save-ClipboardImageMD.ps1
# Created by Bitpusher/The Digital Fox
# v1.7 last updated 2024-05-02
# Script to handle clipboard operations for Markdown files:
# * If there is an image in the clipboard, save it to specific folder and copy Markdown-formatted 
# image link to clipboard for insertion into your document.
# * If the clipboard contains the text of a valid path on the file system to an image file
# copy the indicated image to the folder and create a Markdown-formatted link to it.
# * If the clipboard contains a data object which includes a valid path to an image file
# (a copy to the clipboard was made of an image file through Explorer) copy the indicated 
# image to the folder and create a Markdown-formatted link to it.
#
# Usage:
# powershell -executionpolicy bypass -f .\Save-ClipboardImageMD.ps1 -ImageFolderPath "C:\Users\Username\Documents\Markdown" -RelativeImageFolderPath ".\Images"
#
# Update the "ImageFolderPath" parameter to point to your Markdown image repository location.
# Update the "RelativeImageFolderPath" parameter to be the relative path of the image repository for all your Markdown files.
#
# Needs to be run under PowerShell 5.1 ("Windows PowerShell) - PowerShell 6+ is unable to retrieve images from clipboard - only text
#
#comp #clipboard #markdown #paste #png #jpg #jpeg #gif #image #ahk #macro #powershell #script

#Requires -Version 5.1

param (
 [string]$ImageFolderPath = "$env:USERPROFILE\Documents",
 [string]$RelativeImageFolderPath = ".\images"
)

Add-Type -AssemblyName System.Windows.Forms
$date = Get-Date -Format "yyyyMMddHHmmss"
$clipboard = [System.Windows.Forms.Clipboard]::GetDataObject()

if ($clipboard.ContainsImage()) {
    $filename="$ImageFolderPath\$($date).png"         
    [System.Drawing.Bitmap]$clipboard.getimage().Save($filename, [System.Drawing.Imaging.ImageFormat]::Png)
    Set-Clipboard -Value "![$date.png]($RelativeImageFolderPath\$($date).png)"
} elseif ($clipboard.ContainsText()) {
    $text = $clipboard.GetText()
    # $fileExt = [IO.Path]::GetExtension("$clipboard.gettext")
    $filePath = Get-Item "$text"
    $fileName = $filePath.BaseName
    $fileExt = $filePath.Extension
    if ($fileExt -eq ".jpg" -or $fileExt -eq ".jpeg" -or $fileExt -eq ".png" -or $fileExt -eq ".gif") {
        Copy-Item -Path "$filePath" -Destination "$ImageFolderPath\$fileName-$($date)$fileExt"
        Set-Clipboard -Value "![$fileName-$($date)$fileExt]($RelativeImageFolderPath\$fileName-$($date)$fileExt)"
    } else {
        Set-Clipboard -Value "Clipboard did not contain a valid text image path."
    }
} elseif ($clipboard.ContainsFileDropList()) {
    # $clipboard.GetFileDropList() | out-file "V:\Dropbox\txt\images\bla"
    $filePathText = $clipboard.GetFileDropList()
    $fileLocation = Get-Item $filePathText
    $filePath = $fileLocation.FullName
    $fileName = $fileLocation.BaseName
    $fileExt = $fileLocation.Extension
    if ($fileExt -eq ".jpg" -or $fileExt -eq ".jpeg" -or $fileExt -eq ".png" -or $fileExt -eq ".gif") {
        Copy-Item -Path "$filePath" -Destination "$ImageFolderPath\$fileName-$($date)$fileExt"
        Set-Clipboard -Value "![$fileName-$($date)$fileExt]($RelativeImageFolderPath\$fileName-$($date)$fileExt)"
    } else {
        Set-Clipboard -Value "Clipboard did not contain compatible image file from file system."
    }
} else {
    Set-Clipboard -Value "Clipboard did not contain an image, is not a compatible image file from the file system, and is not a text path to an image file on the file system."
}