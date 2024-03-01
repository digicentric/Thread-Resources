# switches
param(
    [switch]$Update,
	[switch]$DCI,
    [switch]$MEC,
    [switch]$OBEI,
    [switch]$ASI,
    [switch]$RSI,
    [switch]$SCD,
    [switch]$AHMD,
    [switch]$FEMC,
    [switch]$SHSMDI,
    [switch]$OMG,
    [switch]$MVMM,
    [switch]$LA,
    [switch]$IAI,
    [switch]$SMH,
    [switch]$MMS,
    [switch]$THG
)
$switch = @{
    "DCI" = "70cdc36d-2a56-4747-af09-c5cfe3287c63"
    "MEC" = "5efa40dd-7bd0-4464-9bf2-8ffe2f408cd4"
    "OBEI" = "1eabe4db-e802-4001-a585-44c9d454ac9e"
    "AHMD" = "0f59cda6-bbfa-49d8-9120-5344230bd802"
    "RSI" = "ff5c2b93-5d79-419b-b70c-020d7f3b4fc1"
    "ASI" = "68be12ba-3493-4a7d-89dc-4e2126b47717"
    "SCD" = "4fcd18c8-6cb2-49ee-a5ff-70bb1a2c73d6"
    "FEMC" = "2465d5fe-6187-4c5e-81ac-1c12098dfe33"
    "SHSMDI" = "6396eedb-10b2-48ba-a3c6-b20674771ba9"
    "OMG" = "63e52278-53e6-4372-b209-b6bcb37b2753"
    "MVMM" = "ece65017-d756-4d13-a49b-e672f8d143dc"
    "LA" = "f1ddd002-038d-422b-b727-0594dbc96877"
    "IAI" = "992e0704-2aa8-4c75-be0f-4e9af718de03"
    "SMH" = "1a354fa7-8354-43b4-9f0f-fe91a5bc00f4"
    "MMS" = "119877b9-acd5-4530-bb3e-678e08764a5d"
    "THG" = "2940f3c9-0bad-4000-82f3-bdd2da6002c2"
}

# thread install function
function installThread {
    param (
        [string]$appID,
		[switch]$update
    )
	if ((Test-Path "C:\Program Files\Messenger\Messenger.exe") -and (-not $update)) {
		Write-Verbose "Messenger is already installed" -Verbose
	} elseif (Test-Path "C:\Program Files\Messenger\Messenger.exe" -and $update) {
		Write-Verbose "Update parameter has been passed" -Verbose
		try {
			Write-Verbose "Attempting to stop Messenger process" -Verbose
			Stop-Process -Name Messenger -Force
			Write-Verbose "Stopped Messenger process" -Verbose
			try {
				Write-Verbose "Attempting to reinstall Thread for $appID" -Verbose
				Start-Process -FilePath "msiexec" -ArgumentList "/i", "https://assets.getthread.com/messenger/downloads/desktop/messenger.msi", "APP_ID=$appID", "FLOW=customer", "MSIINSTALLPERUSER=", "/qn", "/norestart" -Wait
				Write-Verbose "Reinstalled Thread for $appID" -Verbose
			} catch {
			Write-Error "Error reinstalling Thread: $_"
			}
		} catch {
			Write-Error "Error stopping Messenger process $_"
		}
	} else {
		try {
            Write-Verbose "Attempting to install Thread for $appID" -Verbose
			Start-Process -FilePath "msiexec" -ArgumentList "/i", "https://assets.getthread.com/messenger/downloads/desktop/messenger.msi", "APP_ID=$appID", "FLOW=customer", "MSIINSTALLPERUSER=", "/qn" -Wait
		} catch {
			Write-Error "Error installing Thread: $_"
			Exit 1
		}
	}
}

# icon creation function
function createIcon {
    param (
        [string]$fileName,
        [string]$selectedSwitch
    )
    if (Test-Path "$env:Public\Desktop\$fileName.lnk") {
        Write-Verbose "Custom icon was found" -Verbose
        try {
            Remove-Item -Path "$env:Public\Desktop\$fileName.lnk" -Force
            Write-Verbose "Removed old custom icon" -Verbose
        } catch { 
            Write-Error "Error removing custom icon: $_"
        }
    } else {
        Write-Verbose "No custom icon was found" -Verbose
    }
    try {
		
		Invoke-WebRequest -Uri "https://raw.githubusercontent.com/digicentric/Thread-Resources/main/logos/$selectedSwitch/$selectedSwitch.ico" -OutFile "C:\Program Files\Messenger\$selectedSwitch.ico"
        Write-Verbose "Successfully downloaded $selectedSwitch.ico" -Verbose
    } catch {
        Write-Error "Error downloading icon :$_"
    }
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut("$env:Public\Desktop\$fileName.lnk")
        $Shortcut.IconLocation = "C:\Program Files\Messenger\$selectedSwitch.ico"
        $Shortcut.TargetPath = "C:\Program Files\Messenger\Messenger.exe\"
		$Shortcut.WorkingDirectory = "C:\Program Files\Messenger\"
		$Shortcut.Arguments = "chatgenie://?open=true"
        $Shortcut.Save()
        Write-Verbose "Created custom shortcut for $fileName" -Verbose
    } catch {
        Write-Error "Error creating custom icon: $_"
    }
}

# check for valid switch function & install thread
$selectedSwitch = $null
Write-Verbose "Checking for valid switch key pair" -Verbose
foreach ($switchKey in $switch.Keys) {
    Write-Verbose "Checking switch key: $switchKey" -Verbose
	if ($MyInvocation.BoundParameters.ContainsKey($switchKey)) {
		$selectedSwitch = $switchKey
		break
	}
}
if (-not $selectedSwitch) {
	Write-Error "Invalid or no switch provided. Please pass a valid switch."
	Exit 3
} else {
    Write-Verbose "Found key for $selectedSwitch" -Verbose
	installThread -appID="$switch[$selectedSwitch]"
}

# remove default messenger icon
if (Test-Path "$env:Public\Desktop\Messenger.lnk") {
	try {
		Remove-Item -Path "$env:Public\Desktop\Messenger.lnk" -Force
		Write-Verbose "Messenger.lnk has been removed" -Verbose
	} catch {
		Write-Error "Error removing icon: $_"
	}
} else {
		Write-Verbose "Messenger.lnk does not exist" -Verbose
}

# icon function calls
if ($selectedSwitch -eq "MMS") {
	createIcon -fileName "Meridian Support" -selectedSwitch "mms"
} elseif ($selectedSwitch -eq "SMH") {
	createIcon -fileName "Sun Mar Support" -selectedSwitch "smh"
} elseif (-not ($selectedSwitch -eq "MMS" -or $selectedSwitch -eq "SMH") -and $switch.ContainsKey($selectedSwitch)) {
	Write-Verbose "Selected switch is a valid switch but not MMS or SMH" -Verbose
	createIcon -fileName "Digicentric Chat" -selectedSwitch "dci"
} else {
	Write-Error "No icon URL found for the selected switch"
}

Exit 0