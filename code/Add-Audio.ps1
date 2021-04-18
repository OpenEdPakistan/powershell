[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.PowerPoint") # Load PowerPoint dll
$pptx = New-Object -ComObject PowerPoint.Application # Create COM object from PowerPoint application

function Add-Audio([Microsoft.Office.Interop.PowerPoint.PresentationClass]$Presentation, [string]$Folder)
{

	Set-ExecutionPolicy -executionpolicy bypass
	for ($i = 1; $i -le $Presentation.Slides.Count; $i++) # For each slide in presentation
	{
        $fileName = $i # Initialize file name
        if ($i -le 9)  # If current slide number is a single-digit number
        {
            $fileName = "0" + $i # Add a zero in the file name
        }
        $Slide = [string]$Presentation.Slides($i).SlideShowTransition.SoundEffect.ImportFromFile($Folder + $fileName + ".wav") # Add audio file to slide transition
	}

    Write-Output "Done." # Display success message
}

Add-Audio -Presentation $pptx.ActivePresentation -Folder "C:\Folder-Name\" #Execute Function with parameters
