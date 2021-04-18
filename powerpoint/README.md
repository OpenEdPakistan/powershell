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
[../files/PowerShell-PowerPoint-STD.png]
*The process comprises the following 3 steps, as shown above:*
#### 1. Create Presentation
#### 2. Open Presentation File
#### 3. Execute PowerShell Script
