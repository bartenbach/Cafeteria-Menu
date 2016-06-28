#
#  Download Children's hospital menu and save to desktop
#     By Blake Bartenbach
#
function CheckPSVersion() {
  if ($PSVersionTable.PSVersion.Major -lt 3.0) {
    Write-Host "Your PowerShell is incapable of this feature.  Update it." -ForegroundColor "Yellow"
    Write-Host "https://www.microsoft.com/en-us/download/confirmation.aspx?id=50395" -ForegroundColor "Blue"
  }
}

function Download-Menu() {
	$url = "http://myportal.chsomaha.org/documents/departments/food%20service/cafeteria%20menu.doc"
	$output = "$HOME\Menu.doc"
	Invoke-WebRequest -Uri $url -OutFile $output
}

function main() {
  Write-Host "Checking PS Version..." -ForegroundColor "Green"
  CheckPSVersion
  Write-Host "Downloading Menu..." -ForegroundColor "Green"
  Download-Menu
  Start-Sleep -s 2
  Write-Host "Launching..." -ForegroundColor "Green"
  & 'C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.EXE' "$HOME\Menu.doc"
  Write-Host "Waiting for process to exit..." -ForegroundColor "Green"
  Wait-Process WINWORD
  Write-Host "Removing Menu from Desktop..." -ForegroundColor "Green"
  rm "$HOME\Menu.doc"
  Write-Host "Done!" -ForegroundColor "Green"
}

main