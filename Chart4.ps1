Start-Transcript
# Check Folder Exists
If (Test-Path -Path $env:APPDATA\Microsoft\Templates\Charts) {
    #Get the Folder object
    $Folder = Get-Item -Path $env:APPDATA\Microsoft\Templates\Charts
    Write-Host "Folder already exists." -f Yellow
} Else {
    #Create a New Folder   
    $Folder = New-Item -ItemType Directory -Path $env:APPDATA\Microsoft\Templates\Charts
    Write-Host "Folder Created Successfully!." -f Green


#gets currently logged in user
$Username = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]

Copy-Item -Path "Charts\*" -Destination "C:\Users\$Username\AppData\Roaming\Microsoft\Templates\Charts" -Recurse -Force


New-Item -Path "C:" -Name "Charts" -ItemType "directory" -Force
Copy-Item -Path "Charts\*" -Destination "C:\Users\$Username\AppData\Roaming\Microsoft\Templates\Charts" -Recurse -Force
}
Stop-Transcript
