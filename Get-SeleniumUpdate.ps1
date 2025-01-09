#Checks the Chrome version by registry, chromedriver binary by getting process output,
# then downloads new version if necessary

<#
    # To set a Scheduled task action to run the updater before starting the search:
powershell.exe
-Window Minimized $myprocess = Start-Process "powershell" -ArgumentList "C:\Selenium\Get-SeleniumUpdate.ps1" -PassThru; $myprocess.WaitForExit(); sleep 30; Import-Module Selenium -Global; Sleep 5; & c:\Selenium\Selenium-DoThingsScript.ps1#>

<#
2023-08-24
https://googlechromelabs.github.io/chrome-for-testing/
New site for driver download. Need to download the one that just says chromedriver.

IMPORTANT!      The script will default to downloading a version based on detected local version of chrome.
                There are two functions to get the download link - LatestStable as declared by google, and InstalledVersion of chrome.
                It is hard set to run the 
----------------------------------------------------------------------------
2023-12-15
Fixed failure from missing chromedriver.exe by using WebDriver.dll as the locator file for the module.

#>


Function Get-SeleniumUpdate {
    $Script:InstalledChromeVersion = $null
    $Script:BinaryVersion = $null
    $Script:BinaryMajorVersion = $null
    $Script:InstalledChromeVersion = $null
    $Script:InstalledChromeMajorVersion = $null
    $Script:SeleniumModuleBinaryFile = $null
    $Script:SeleniumModuleBinaryFolder = $null
    $UpdateSelenium = $null
    $DownloadLink = $null
    Function Get-ChromeDownloadForLatestStableVersion {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        #$StableUri = 'https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE'
        #$StableVersion = (iwr -uri $StableUri).ToString()
        $JsonUri = 'https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json'
        $Json = (Invoke-RestMethod -Uri $JsonUri -Method GET)
        $DownloadLink = ($Json.Channels.Stable | ? Channel -eq "Stable").Downloads.chromedriver | ? platform -eq "win64" | select -expand url
        Return $DownloadLink
    }
    Function Get-ChromeDownloadForInstalledVersion {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $Script:InstalledChromeVersion = (Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo.ProductVersion.Trim()
        $GetVersion = $null; 
        For ($i=0; $i -lt ($Script:InstalledChromeVersion.Split('.').Count -1); $i++) { $GetVersion = $GetVersion + "$($Script:InstalledChromeVersion.Split('.')[$i])`." }
        $JsonUri = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'
        $Json = (Invoke-RestMethod -Uri $JsonUri -Method GET)
        $DownloadLink = ($Json.Versions | ? Version -match $GetVersion).Downloads.chromedriver | ? Platform -eq "win64" | select -Expand Url -Last 1
        Return $DownloadLink
    }
    function Get-ProcessOutput {
        #For retrieving the output from chromedriver.exe --version
        Param ( [Parameter(Mandatory=$true)]$FileName, $Args )
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        $process.StartInfo.FileName = $FileName
        if($Args) { $process.StartInfo.Arguments = $Args }
        $out = $process.Start()
        $StandardError = $process.StandardError.ReadToEnd()
        $StandardOutput = $process.StandardOutput.ReadToEnd()
        $output = New-Object PSObject
        $output | Add-Member -type NoteProperty -name StandardOutput -Value $StandardOutput
        $output | Add-Member -type NoteProperty -name StandardError -Value $StandardError
        return $output
    }
    Function Get-SeleniumAndChromeVersions {
        #Determine the local chromedriver binary major version
        if (Test-Path $Script:SeleniumModuleBinaryFile -ErrorAction SilentlyContinue) {
            $Script:BinaryVersion = (Get-ProcessOutput -FileName "$Script:SeleniumModuleBinaryFile" '--version').StandardOutput.Split(" ")[1]
            $Script:BinaryMajorVersion = $null; 
            For ($i=0; $i -lt ($Script:BinaryVersion.Split('.').Count -1); $i++) { $Script:BinaryMajorVersion = $Script:BinaryMajorVersion + "$($Script:BinaryVersion.Split('.')[$i])`." }
        } Else { $Script:UpdateSelenium = $False; Write-Host -f Red "`n   Get-SeleniumUpdate: Not found: $Script:SeleniumModuleBinaryFile`n" }
        
        #Determine the installed chrome major version 
        $Script:InstalledChromeVersion = (Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo.ProductVersion.Trim()
        $Script:InstalledChromeMajorVersion = $null; 
        For ($i=0; $i -lt ($Script:InstalledChromeVersion.Split('.').Count -1); $i++) { $Script:InstalledChromeMajorVersion = $Script:InstalledChromeMajorVersion + "$($Script:InstalledChromeVersion.Split('.')[$i])`." }
        $Script:InstalledChromeMajorVersion = $Script:InstalledChromeMajorVersion
    }
    
    #For working with selenium binaries
    $PossiblePaths = Join-Path ($env:PSModulePath.Split(';') -Match "WindowsPowerShell\\Modules") ("\Selenium\assemblies\WebDriver.dll") | Select -Unique
    $Script:SeleniumModuleBinaryFile = $PossiblePaths[(Test-Path -Path $PossiblePaths).indexOf($true)]
    if (Test-Path $Script:SeleniumModuleBinaryFile -ErrorA Si) {write-host -f green "    `nSelenium module WebDriver.dll found: $Script:SeleniumModuleBinaryFile `n"}
    $Script:SeleniumModuleBinaryFile   = Join-Path (get-item $Script:SeleniumModuleBinaryFile | Split-Path -Parent) "chromedriver.exe"
    $Script:SeleniumModuleBinaryFolder = Split-Path $Script:SeleniumModuleBinaryFile -Parent

    #Break if no binaries are found.
    if ((Test-Path $SeleniumModuleBinaryFolder -ErrorA Si) -eq $False) {write-host -f Red "`n    Did not find a Selenium driver binaries in any modules paths:"; $PossiblePaths | % {$_}; "`n"; BREAK}

    Get-SeleniumAndChromeVersions

    #Determine if $Script:UpdateSelenium = $true/$false
    if ($InstalledChromeMajorVersion -eq $BinaryMajorVersion) { $UpdateSelenium = $False } Else { $UpdateSelenium = $True  }
    if ($UpdateSelenium -eq $False) { write-host -f green "`n    Selenium and Chrome binary versions match. `n"; BREAK } #Break if versions match

    #Kill any chromedrivers so we can replace with new version, then make a backup and remove old version of sel bin
    If (Get-Process | Where ProcessName -match chromedriver) {Get-Process | Where ProcessName -match chromedriver | Stop-Process -Force}
    
    #   !!!!!!      Get the link for the INSTALLED major version of chrome, not for the Latest Stable as declared by google.
    $DownloadLink = Get-ChromeDownloadForInstalledVersion 
    
    $BinaryZipDownloadPath = "$env:USERPROFILE\Downloads\chromedriver_win64.zip"
    if (test-path $BinaryZipDownloadPath -ErrorA Si)  {rm $BinaryZipDownloadPath -force}
    Start-BitsTransfer -Source $DownloadLink -Destination $BinaryZipDownloadPath
    Sleep 2
    
    #unzip the selenium download and unblock the new binary
    if (Test-Path $BinaryZipDownloadPath) {
        Unblock-File -Path $BinaryZipDownloadPath
        Expand-Archive -Path $BinaryZipDownloadPath -DestinationPath $SeleniumModuleBinaryFolder -Force
    } Else { 
        write-host -f Red "  Tried to unzip    $BinaryZipDownloadPath    . Not found. `n    Nothing to do. Breaking. `n"; 
        break 
    }
    
    gci (Join-Path $SeleniumModuleBinaryFolder "chromedriver-win64") | Move-Item -Destination $SeleniumModuleBinaryFolder -Force
    rm (Join-Path $SeleniumModuleBinaryFolder "chromedriver-win64") -recurse -force
    
    Unblock-File -Path $SeleniumModuleBinaryFolder\chromedriver.exe

    Get-SeleniumAndChromeVersions

    write-host -f Green "`n               Selenium update was attempted. Result:"
    Write-Host -f Cyan "    Selenium: " $Script:BinaryVersion
    Write-Host -f Cyan "      Chrome: " $Script:InstalledChromeVersion "`n"

}

Get-SeleniumUpdate
