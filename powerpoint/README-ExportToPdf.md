# Exporting Microsoft PowerPoint files to PDF using PowerShell
## Why?
To avoid manually exporting Microsoft PowerPoint presentations to Adobe Acrobat format (PDF).
## What?
PowerShell script to perform the same task automatically.
## How?
### Pre-requisites
*The following two are pre-requisites for executing the PowerShell script file.*
#### 1. PowerShell IDE
Open the PowerShell file _Export-PptxToPdf.ps1_ in Microsoft Visual Studio Code, Windows PowerShell ISE, or PowerShell Studio 2021. I used Windows PowerShell ISE on my machine.
#### 2. Execution Policy
If you get an error that the file is not signed when you run the script, you have to execute the following PowerShell command in your environment.
```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
*Note that the above command only affects your current session.*
### Steps
*Execute the PowerShell code given below:*
```
function Export-PptxToPdf($inputFile)
{
	[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
	[Reflection.Assembly]::LoadWithPartialname("Office") > $null

	$msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue

	$ppFixedFormatIntentScreen = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentScreen # Intent is to view exported file on screen.
	$ppFixedFormatIntentPrint =  [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint  # Intent is to print exported file.

	$ppFixedFormatTypeXPS = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypeXPS  # XPS format
	$ppFixedFormatTypePDF = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF  # PDF format

	$ppPrintHandoutVerticalFirst = 1   # Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it.
	$ppPrintHandoutHorizontalFirst = 2 # Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it.

	$ppPrintOutputSlides = 1               # Slides
	$ppPrintOutputTwoSlideHandouts = 2     # Two Slide Handouts
	$ppPrintOutputThreeSlideHandouts = 3   # Three Slide Handouts
	$ppPrintOutputSixSlideHandouts = 4     # Six Slide Handouts
	$ppPrintOutputNotesPages = 5           # Notes Pages
	$ppPrintOutputOutline = 6              # Outline
	$ppPrintOutputBuildSlides = 7          # Build Slides
	$ppPrintOutputFourSlideHandouts = 8    # Four Slide Handouts
	$ppPrintOutputNineSlideHandouts = 9    # Nine Slide Handouts
	$ppPrintOutputOneSlideHandouts = 10    # Single Slide Handouts

	$ppPrintAll = 1            # Print all slides in the presentation.
	$ppPrintSelection = 2      # Print a selection of slides.
	$ppPrintCurrent = 3        # Print the current slide from the presentation.
	$ppPrintSlideRange = 4     # Print a range of slides.
	$ppPrintNamedSlideShow = 5 # Print a named slideshow.

	$ppShowAll = 1             # Show all.
	$ppShowNamedSlideShow = 3  # Show named slideshow.
	$ppShowSlideRange = 2      # Show slide range.

	$application = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

	$inputFile = Resolve-Path $inputFile
	$outputFile = [System.IO.Path]::ChangeExtension($inputFile, ".pdf")
	
	$application.Visible = $msoTrue
	$presentation = $application.Presentations.Open($inputFile, $msoTrue, $msoFalse, $msoFalse)
	$printOptions = $presentation.PrintOptions
	$range = $printOptions.Ranges.Add(1,$presentation.Slides.Count) 
	$printOptions.RangeType = $ppShowAll
	
	$presentation.ExportAsFixedFormat($outputFile, $ppFixedFormatTypePDF, $ppFixedFormatIntentScreen, $msoTrue, $ppPrintHandoutHorizontalFirst, $ppPrintOutputSlides, $msoFalse, $range, $ppPrintAll, "Slideshow Name", $False, $False, $False, $False, $False)
	
	$presentation.Close()
	$presentation = $null
	
	if($application.Windows.Count -eq 0)
	{
		$application.Quit()
	}
	
	$application = $null
	
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();

    Write-Output "Done." # Display success message
}

Export-PptxToPdf -inputFile "C:\FolderName\SubfolderName\FileName.pptx"
```
View the PowerShell file here: https://github.com/OpenEdPakistan/powershell/blob/main/code/Export-PptxToPdf.ps1
<br /><br />
*Originally published [here](https://gist.github.com/ap0llo/05cef76e3c4040ee924c4cfeef3f0b40#file-export-presentation-ps1)*
