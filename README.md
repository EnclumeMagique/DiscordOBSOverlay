**NOTE:**

Much is hardcoded, if you want to use this you will likely have to change the code a bit, this is why no exe is included

-------------

**REQUIREMENTS:**

pywinauto

obs-websocket-py

Flask

waitress

comtypes

tkinter

You can use pip to install all of those

------------------------
**Setup:**

Muted icon on OBS (Default name: Discord Mute)

Deafened icon on OBS (Default name: Discord Deafen)

Discord logo or any indication to the user that it is referring to the Discord mic on OBS (Default name: Discord audio info)

Error when not connected to the overlay on OBS (Default name: Discord audio info error)

Talking icon on OBS when inverse mode is on (Default name: Discord inverse unmute)

------------------------

**IF YOU HAVE A STREAMDECK OR ANY PROGRAMABLE BUTTON SYSTEM:**

**These are powershell scripts you can run to toggle certain actions out of the EXE**

# ToggleOverlay.ps1
$command = @{
    command = "toggle_overlay"
}
$commandJson = $command | ConvertTo-Json
Invoke-WebRequest -Uri "http://localhost:5001/toggle_overlay" -Method POST -Body $commandJson -ContentType "application/json"
Reverse mode On/Off:
# ToggleInverseMode.ps1
$command = @{
    command = "toggle_inverse"
}
$commandJson = $command | ConvertTo-Json
Invoke-WebRequest -Uri "http://localhost:5001/toggle_inverse" -Method POST -Body $commandJson -ContentType "application/json"
Reverse force off:
# ForceReverseOff.ps1
$command = @{
    command = "force_reverse_off"
}
$commandJson = $command | ConvertTo-Json
Invoke-WebRequest -Uri "http://localhost:5001/force_reverse_off" -Method POST -Body $commandJson -ContentType "application/json"
Reverse force On:
# ForceReverseMode.ps1
$command = @{
    command = "force_reverse"
}
$commandJson = $command | ConvertTo-Json
Invoke-WebRequest -Uri "http://localhost:5001/force_reverse" -Method POST -Body $commandJson -ContentType "application/json"
