# Automating Microsoft PowerPoint using PowerShell
## Why?
To avoid manually inserting Urdu Audio files (with the .wav extension) into Microsoft PowerPoint presentation.
## What?
PowerShell script to perform the same task automatically in seconds.
## How?
### Pre-requisites
*The following two are pre-requisites for executing the PowerShell script file.*
#### 1. PowerShell IDE
Open the PowerShell file _Add-Audio.ps1_ in Microsoft Visual Studio Code, Windows PowerShell ISE, or PowerShell Studio 2021. I used Windows PowerShell ISE on my machine.
#### 2. Execution Policy
If you get an error that the file is not signed when you run the script, you have to execute the following PowerShell command in your environment.
```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
*Note that the above command only affects your current session.*
### Steps
![PowerPoiny Automation](../files/PowerShell-PowerPoint-STD.png)
*The process comprises the following 3 steps, as shown above:*
#### 1. Create Content
a. The User enter Urdu text into the text-to-speech system.
b. The generated Urdu audio files are stored in a Folder.
c. The educational content is created and stored in a PowerPoint presentation.
*At this stage, the audio and visual parts exist separately.*
#### 2. Open Presentation File
Run the PowerPoint application as an administor on the computer, and open the presentation file that the audio needs to be added to (created in step 1c)
#### 3. Execute PowerShell Script
a. The User executes the PowerShell code in the script file.
b. The narration in the audio files is embedded into the presentation. This is achieved by naming the audio files as numbers, eac corresponding to a slide number.
c. The PowerPoint slides are updated by the executed PowerShell script.
d. The slides with the audio and visual content both are ready for use by the User. Keep in mind that the file needs to be saved manually (either by using the menu option or the shortcut-key) in order for the updates to be committed to the disk.
