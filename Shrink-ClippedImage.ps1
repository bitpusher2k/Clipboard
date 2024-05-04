#           Bitpusher
#            \`._,'/
#            (_- -_)
#              \o/
#          The Digital
#              Fox
#    https://theTechRelay.com
# https://github.com/bitpusher2k
#
# Shrink-ClippedImage.ps1
# Created by Bitpusher/The Digital Fox
# v1.1 last updated 2024-05-02
# Shrinks an image in the Windows clipboard & copies it back into clipboard resized.
#
# Usage:
# powershell -executionpolicy bypass -f .\Shrink-ClippedImage.ps1 -Ratio 60
#
# Set the "Ratio" to the percentage of original to resize image.
#
# Run with an image in the clipboard.

#powershell #script #clipboard #image #resize #shrink #ahk #macro

param (
 [string]$Ratio = 60 # Shriink/Expand ratio: Shrink image to 60% of original in this case
)

Add-Type -AssemblyName System.Windows.Forms
[void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$cbImage = [System.Windows.Forms.Clipboard]::GetImage()
If ($cbImage -ne $null) {
  $Ratio = $Ratio / 100
  $newWidth = [int]($cbImage.Width * $Ratio)
  $newHeight = [int]($cbImage.Height * $Ratio)
  $bmImage = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
  $newImage = [System.Drawing.Graphics]::FromImage($bmImage)
  $newImage.DrawImage($cbImage, $(New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $newWidth, $newHeight))
  [Windows.Forms.Clipboard]::SetImage($bmImage)
  $cbImage.Dispose()
  $bmImage.Dispose()
  $newImage.Dispose()
}
