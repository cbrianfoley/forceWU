'''INSTALL WINDOWS UPDATES'''

#Log all the things to C:\Windowsupdate.log and continue even if there are errors
$ErrorActionPreference = "SilentlyContinue" 
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -path C:\WindowsUpdate.log -Append | Out-Null
If ($Error) {
    $Error.Clear()
    }

#Make the update session and search for updates
$UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl 
$Searcher = New-Object -ComObject Microsoft.Update.Searcher 
$Session = New-Object -ComObject Microsoft.Update.Session
Write-Output "Initializing and Checking for Applicable Updates. Please wait ..." 
$Result = $Searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

#Print the result of the search
If ($Result.Updates.Count -EQ 0) {
    Write-Output "There are no applicable updates for this computer." 
    } Else { 
    Write-Output "Preparing List of Applicable Updates For This Computer ..."
    For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) {
        $DisplayCount = $Counter + 1 
        $Update = $Result.Updates.Item($Counter) 
        $UpdateTitle = $Update.Title 
        Write-Output "$DisplayCount -- $UpdateTitle" 
        }
    $Counter = 0 
    $DisplayCount = 0

#Download the updates
    Write-Output "Initializing Download of Applicable Updates ..." 
    $Downloader = $Session.CreateUpdateDownloader() 
    $UpdatesList = $Result.Updates 
    For ($Counter = 0; $Counter -LT $Result.Updates.Count; $Counter++) { 
        $UpdateCollection.Add($UpdatesList.Item($Counter)) | Out-Null 
        $ShowThis = $UpdatesList.Item($Counter).Title 
        $DisplayCount = $Counter + 1 
        Write-Output "$DisplayCount -- Downloading Update $ShowThisr" 
        $Downloader.Updates = $UpdateCollection 
        $Track = $Downloader.Download() 
        If (($Track.HResult -EQ 0) -AND ($Track.ResultCode -EQ 2)) { 
            Write-Output "Download Status: SUCCESS" 
            } Else { 
            Write-Output "Download Status: FAILED With Error -- $Error()" 
            $Error.Clear() 
            Write-Output 
            } 
        } 
    $Counter = 0 
    $DisplayCount = 0

#Install the updates
    Write-Output "Starting Installation of Downloaded Updates ..." 
    $Installer = New-Object -ComObject Microsoft.Update.Installer 
    For ($Counter = 0; $Counter -LT $UpdateCollection.Count; $Counter++) {
        $Track = $Null
        $DisplayCount = $Counter + 1
        $WriteThis = $UpdateCollection.Item($Counter).Title
        Write-Output "$DisplayCount -- Installing Update: $WriteThis" 
        $Installer.Updates = $UpdateCollection 
        Try { 
            $Track = $Installer.Install() 
            Write-Output "Update Installation Status: SUCCESS" 
            } Catch { 
            [System.Exception] 
            Write-Output "Update Installation Status: FAILED With Error -- $Error()" 
            $Error.Clear() 
            } 
        } 
    }
Stop-Transcript