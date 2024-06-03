#Products Path
$ProductsPath = "HKLM:\SOFTWARE\Classes\Installer\Products\"
$FeaturesPath = "HKLM:\SOFTWARE\Classes\Installer\Features\"
$UpgradesPath = "HKLM:\SOFTWARE\Classes\Installer\UpgradeCodes\"
$UpgradesPath2 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\"
$UninstallsPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$DefaultUserInstallPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*\InstallProperties\"
$OtherPath = "HKLM:\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed\"
$UninstallsPath2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"

$NinjaRegPath = "HKLM:\SOFTWARE\NinjaRMM LLC"
$NinjaRegPath2 = "HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC"
$NinjaProgramDataPath = "C:\ProgramData\NinjaRMMAgent\"

Stop-Process -Name "NinjaRMMAgent" -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Name "NinjaRMMAgentPatcher" -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Name "njbar" -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Name "NinjaRMMProxyProcess64" -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Name "nmsmanager" -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Name "lockhart" -Force -ErrorAction SilentlyContinue | Out-Null

$RegKeys = Get-ChildItem -Path $ProductsPath
Foreach ($RegKey in $RegKeys) {
      $Object = Get-ItemProperty Registry::$RegKey
      if ($Object."ProductName" -eq "NinjaRMMAgent") {
            $ProductPath = $ProductsPath + $Object."PSChildName"
            $FeaturePath = $FeaturesPath + $Object."PSChildName"
            $UpgradePath = $UpgradesPath + $Object."PSChildName"
            $UpgradePath2 = $UpgradesPath2 + $Object."PSChildName"

            Remove-Item -Path $ProductPath -Recurse -ErrorAction SilentlyContinue | Out-Null
            Remove-Item -Path $FeaturePath -Recurse -ErrorAction SilentlyContinue | Out-Null
            Remove-Item -Path $UpgradePath -Recurse -ErrorAction SilentlyContinue | Out-Null
            Remove-Item -Path $UpgradePath2 -Recurse -ErrorAction SilentlyContinue | Out-Null
    }
}

$RegKeys = Get-ChildItem -Path $UninstallsPath
Foreach ($RegKey in $RegKeys) {
      $Object = Get-ItemProperty Registry::$RegKey
      if ($Object.DisplayName -eq "NinjaRMMAgent") {
        $UninstallPath = $UninstallsPath + $Object."PSChildName"
        Remove-Item -Path $UninstallPath -Recurse -Force
    }
}

$RegKeys = Get-ChildItem -Path $UninstallsPath2
Foreach ($RegKey in $RegKeys) {
      $Object = Get-ItemProperty Registry::$RegKey
      if ($Object.DisplayName -eq "NinjaRMMAgent") {
        $UninstallPath = $UninstallsPath2 + $Object."PSChildName"
        Remove-Item -Path $UninstallPath -Recurse -Force
    }
}

$RegKeys = Get-ChildItem -Path $DefaultUserInstallPath
Foreach ($RegKey in $RegKeys) {
      $Object = Get-ItemProperty Registry::$RegKey
      if ($Object.DisplayName -eq "NinjaRMMAgent") {
        $UninstallPath = $Object.PSParentPath
        Remove-Item -Path $UninstallPath -Recurse -Force
    }
}

$RegKeys = Get-ChildItem -Path $OtherPath
Foreach ($RegKey in $RegKeys) {
      $Object = Get-ItemProperty Registry::$RegKey
      if ($Object.PSChildName -Match "NinjaRMMAgent") {
        $UninstallPath = $Object.PSPath
        Remove-Item -Path $UninstallPath -Recurse -Force
    }
}

$ProgramFiles = Get-ChildItem -Path "C:\Program Files (x86)\"
Foreach ($ProgramFolder in $ProgramFiles) {
    $FileName = "NinjaRMMAgent.exe"
    if(Test-Path -Path "$($ProgramFolder.FullName)\$($FileName)") {
        Remove-Item -Path $ProgramFolder.FullName -Recurse -Force
    }
}

Remove-Item -Path $NinjaRegPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path $NinjaRegPath2 -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path $NinjaProgramDataPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
cmd.exe /c "sc delete NinjaRMMAgent" -ErrorAction SilentlyContinue | Out-Null
cmd.exe /c "sc delete lockhart" -ErrorAction SilentlyContinue | Out-Null
cmd.exe /c "sc delete nmsmanager" -ErrorAction SilentlyContinue | Out-Null
