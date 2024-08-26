**NOTES:**

Much is hardcoded, if you want to use this you will likely have to change the code a bit, this is why no exe is included

This is a simple python script, just get a code editor, paste the code, edit it if you want and run it (It works as an exe too, however was only tested with pyinstaller)

To export: python -m PyInstaller --onefile --noconsole overlay.py

No Discord bot needed, just the discord app on windows (may work on other platforms but i dont think so)

In OBS websocket server settings, make sure "Enable WebSocket Server" is ticked, the script will not work otherwise

Script was only tested in Windows 11

You can have Discord however you want but you cannot minimize it, the code will not work if it is.

-------------

**REQUIREMENTS:**

pip install -r requirements.txt

https://github.com/obsproject/obs-websocket

The GUI does not provide the overlay, it just tells it what to do, you need to setup the assets below

------------------------
**Setup:**

Muted icon on OBS (Default name: Discord Mute)

Deafened icon on OBS (Default name: Discord Deafen)

Discord logo or any indication to the user that it is referring to the Discord mic on OBS (Default name: Discord audio info)

Error when not connected to the overlay on OBS (Default name: Discord audio info error)

Talking icon on OBS when inverse mode is on (Default name: Discord inverse unmute)

**These can be text, images, media or website, it just needs to be a source in a singular OBS scene**

------------------------

**IF YOU HAVE A STREAMDECK OR ANY PROGRAMABLE BUTTON SYSTEM:**

**These are powershell scripts you can run to toggle certain actions out of the EXE**

You will need this plugin if you use the streamdeck: https://marketplace.elgato.com/product/windows-scriptdeck-857f01dd-8fd4-44d5-8ec7-67ac850b21d3

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
