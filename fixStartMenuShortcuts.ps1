#   Fix Start Menu Shortcuts
#       Check registry for each item in the $Programs array, and add shortcuts/links to
#       Windows Start Menu for items found.  Script must be run with local administrator
#       permissions.
#
#   $Programs
#       'Shortcut name' = 'Executable name'
#       Several commonly used programs are already on the list for reference.  Feel free
#       to adjust this list as required.
$Programs = @{ 
    'Excel' = 'Excel.exe'
    'Word' = 'Winword.exe'
    'Outlook' = 'OUTLOOK.EXE'
    'Access' ='MSACCESS.EXE'
    'Publisher' = 'MSPUB.EXE'
    'OneNote' = 'OneNote.exe'
    'PowerPoint' = 'powerpnt.exe'
    'Microsoft Edge' = 'msedge.exe'
    'Google Chrome' = 'chrome.exe'
    'Adobe Reader' = 'AcroRd32.exe'
    'VLC Player' = 'vlc.exe'
    'Acrobat Pro' = 'Acrobat.exe'
    'Autodesk Desktop App' = 'AutodeskDesktopApp.exe'
    '7zip File Manager' = '7zFM.exe'
    'LanSchool Teacher' = 'teacher.exe'
}

foreach( $p in $Programs.Keys ){
    $WShell = New-Object -comObject WScript.Shell
    $Shortcut = $WShell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$p.lnk") 
    if (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$($programs.$p)") { 
        $Shortcut.TargetPath = [string](Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$($programs.$p)").'(default)'
        $Shortcut.save()
    }    
}

#   Some apps do not have an "App Path" entry in the registry.
#   These individual apps have their own shortcut creation logic below.
#   'Shortcut Name' = 'Full path to .exe' 

$otherPrograms = @{
    'Airtame' = 'C:\Program Files (x86)\Airtame\Airtame.exe'
    'InTouch' = 'C:\Program Files (x86)\InTouch\InTouchReceiptingSystem.exe'
    'exacqVision Client' = 'C:\Program Files\exacqVision\Client\edvrclient.exe'
}
foreach( $p in $otherPrograms.Keys ){
    $WShell = New-Object -comObject WScript.Shell
    $Shortcut = $WShell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$p.lnk")
    $myVar = $otherPrograms.$p 
    if (Test-Path -Path $myVar -PathType Leaf) { 
        $Shortcut.TargetPath = $myVar
        $Shortcut.save()
    }    
}
