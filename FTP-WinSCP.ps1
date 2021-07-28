[CmdletBinding()] Param(
    [String] $SN_FromJSON,
    [switch] $SN_FromAPI, 
    [switch] $ChooseSN,
    [switch] $CollectionSN,
    [switch] $ExcludeSN,
    [int] $SplitSN,
    [switch] $NoConfig,
    [ValidateSet("CityFit", "Corner", "Fullscreen", "LED", "Testowy_config")]
    [String] $Config,
    [ValidateSet("Blackscreen", "Cityfit", "Corner", "Empik", "SN")]
    [String] $Default,
    [switch] $CheckPlayer = $true,
    [switch] $Release,
    [switch] $ReleaseWebUpdate,
    [switch] $v3_Release,
    [switch] $ForceReleaseUpdate,
    [switch] $Restart,
    [switch] $Restart_Player,
    [switch] $UpdateNetAndRedist,
    [switch] $UpdateAdmin_W10,
    [switch] $updatePowershell_Win8,
    [switch] $updatePowershell_Win7,
    [switch] $CustomFile,
    [switch] $CustomAllFilesFromDirectory,
    [switch] $DeleteFiles,
    [switch] $OldAndTempFilesDelete,
    [switch] $CompleteUpdate,
    [switch] $ScreenshotService,
    [switch] $SSH,
    $InputServers
)

# Default encoding
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

# URL
$requestURL = 'http://***.pl/'
# Headers
$requestHeaders = @{'sntoken' = '***'; 'Content-Type' = 'application/json' }

$ftp = (-join($env:USERPROFILE, "\Desktop\FTP"))
$failed = @()
$notConnected = @()
$done = @()
$time = get-date -Format "dd.MM.yyyy HH-mm"

# Parameters
$builtinParameters = @("ErrorAction","WarningAction","Verbose","ErrorVariable","WarningVariable","OutVariable","OutBuffer","Debug","InformationAction","InformationVariable","PipelineVariable")
$skipParameters = @("SN_FromJSON","SN_FromAPI","ChooseSN","CollectionSN","ExcludeSN", "SplitSN","Initiation","InputServers","CustomFile","CustomAllFilesFromDirectory","DeleteFiles")
$boundParameters = @()

class FileObject {
    [string]$File
    [string]$DestinationPath
}

function FileTransferProgress {
    param(
        $e
    )

    if (($Null -ne $script:lastFileName) -and ($script:lastFileName -ne $e.FileName)) {
        Write-Host 
    }
 
    Write-Host -NoNewline ("`r{0} ({1:P0})" -f ( -join ((Get-date), " [Send]: ", (Split-Path $e.FileName -Leaf))), $e.FileProgress)
    $script:lastFileName = $e.FileName
}

function Log {
    param (
        [string]$msg,
        [switch]$e,
        [switch]$not,
        [switch]$del
    )
    
    if ($e) {
        $message = (-join($(Get-Date), " [ERROR]: ", $msg))
    }
    elseif($not) {
        $message = $msg
    }
    elseif($del) {
        $message = (-join($(Get-Date), " [REMOVE]: ", $msg))
    }
    else {
        $message = (-join($(Get-Date), " [SEND]: ", $msg))
    }

    $message | Out-File -FilePath $global:log -Encoding OEM -Append
}

Function Get-FileName {   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.ShowDialog() | Out-Null
    
    return $OpenFileDialog.filename
} 

Function Get-Directory {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "Desktop"

    if ($foldername.ShowDialog() -eq "OK") {
        $folder += $foldername.SelectedPath
    }
    
    return $folder
}

function GetComputers {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline)]
        [ValidateNotNull()]
        [String]$snNumber
    )
    
    # Body
    $requestBody = @"
{

"names": ["$($snNumber)"]

}
"@

    # Request
    try {
        $request = Invoke-WebRequest -Uri $requestURL -Method POST -Body $requestBody -Headers $requestHeaders -ea Stop
    }
    catch [exception] {
        $Error[0]
        Exit 1
    }

    # Creating PS array of sn
    if ($request.StatusCode -eq 200) {
        $requestContent = $request.content | ConvertFrom-Json
    }
    else {
        Write-host ( -join ("Received bad StatusCode for request: ", $request.StatusCode, " - ", $request.StatusDescription)) -ForegroundColor Red
        Exit 1
    }

    return $requestContent
}

function SendingConfigConfirmation {
    param (
        [string] $version,
        [switch] $not
    )
    
    $cnfgTitle = ' '
    $cnfgMsg = "Czy chcesz kontynuowac wysylanie?"
    $cnfgYes = New-Object System.Management.Automation.Host.ChoiceDescription '&Tak'
    $cnfgNo = New-Object System.Management.Automation.Host.ChoiceDescription '&Nie'
    $cnfgOptions = $cnfgYes, $cnfgNo

    if (!($not)) {
        Write-host " `n "
        Write-Warning "Wersja config.json - $version"
    }
    else {
        Write-host " `n "
        Write-Warning "Nie wybrano wersji pliku config.json"
    }

    $cnfgResposne = $Host.UI.PromptForChoice($cnfgTitle, $cnfgMsg, $cnfgOptions, 1)
    Switch ($cnfgResposne) {
        "0" { return 0 }
        "1" { return 1 }
    }
}

# Clean logs 
if (!(test-path -Path "$ftp\logs")) {
    mkdir "$ftp\logs" | out-null
}

gci $ftp -Filter "Transfer FTP*"  | ? { (New-TimeSpan -Start $_.CreationTime -End (Get-Date -format "yyyy-MM-dd HH:mm:ss")).Days -gt 1 } | % {
    move-item -Path $_.FullName -Destination "$ftp\logs" -Force -Verbose
}

gci "$ftp\logs" | ? { (New-TimeSpan -Start $_.CreationTime -End (Get-Date -format "yyyy-MM-dd HH:mm:ss")).Days -gt 7 } | % {
    Remove-Item -Path $_.FullName -Force -Verbose
}

# log
if($InputServers -eq $null) {
    $logDest = Split-Path $myinvocation.InvocationName -Parent
    $preLog = New-Item -Path $logDest -Name (-join("Transfer FTP - ", $time, ".log")) -ItemType File -Force | select VersionInfo
    $global:log = $preLog.VersionInfo.filename
}
else {
    $inputShort = Split-Path $InputServers -Leaf
    $logDest = Split-Path $InputServers -Parent
    $preLog = New-Item -Path $logDest -Name (-join("Transfer FTP - input ", $inputShort, " - ", $time, ".log")) -ItemType File -Force | select VersionInfo
    $global:log = $preLog.VersionInfo.filename
}

# creating sn list
if (!($SN_FromAPI) -and ($InputServers -eq $null)){
    if (Test-Path $SN_FromJSON) {
        $servers = Get-Content -Raw -Path $SN_FromJSON | ConvertFrom-Json
    }
    else {
        Write-Host "Bledna sciezka do pliku .json!" -ForegroundColor Red -BackgroundColor Black 
        Log -not -msg "Bledna sciezka do pliku .json"
        Exit
    }
}
elseif ($InputServers -ne $null) {
    $servers = Import-Clixml -Path $InputServers
    Remove-Item $InputServers 
    Write-Host "Wysylanie do:"
    Write-Host $servers.name -NoNewline
}
elseif ($SN_FromJSON -ne $null) {
    [string]$snNumbers = Read-Host -Prompt "`nWpisz numery SN przedzielajac je przecinkiem (bez spacji)"

    if ($snNumbers -eq ($null -or "")) {
        "`n"
        Log -not -msg "Brak wpisanych numerow SN"
        Write-host "Wpisz prawidlowe numery SN!" -ForegroundColor Red
        "`n"
        Break
    }

    if ($snNumbers.StartsWith('sn') -eq $false) {
        $snNumbers = 'sn' + $snNumbers
    }

    $toLog = $snNumbers -replace ',', ",sn"
    $snNumbers = $snNumbers -replace ',', "`",`"sn"

    Log -not -msg "`nPobieranie z ScreenNetworkAPI informacji dla komputerow: $toLog`n`n##########################"
    $servers = GetComputers -snNumber $snNumbers

    if($servers.count -eq 0) {
        "`n"
        Log -not -msg "Brak pobranych danych z API"
        Write-host "Brak pobranych danych z API" -ForegroundColor Red
        "`n"
        Break
    }
}

# Config
if ($Config) {
    if ($InputServers -eq $null) {
        if (((SendingConfigConfirmation -version "$Config") -eq 1)) { Break }
    }
}

elseif ($Release -and (!($Config))) {
    if ($InputServers -eq $null) {
        if ((SendingConfigConfirmation -not) -eq 1) { Break }
    }

    $NoConfig = $true
}

# Custom file
if ($CustomFile) {
    "`n"
    Write-Output "Choose file to send"
    Start-Sleep -s 1
    [String]$filePath = Get-FileName
    Write-Output ( -join ("Sending file : ", $filePath))
    [String]$sftpPath = Read-Host -Prompt "Destination path (remote) "
}

if ($CustomAllFilesFromDirectory) {
    "`n"
    Write-Output "Choose directory to send all files from it"
    Start-Sleep -s 1
    [String]$directoryPath = Get-Directory
    Write-Output ( -join ("Sending all files from directory : ", $directoryPath))
    [String]$sftpPath = Read-Host -Prompt "Destination path (remote) "
}

if ($DeleteFiles) {
    "`n"
    [String]$removeFilesPath = Read-Host -Prompt "File / Files path (remote) "
    Log -not -msg "`nUsuwane pliki:`n $removeFilesPath`n"
    Log -not -msg "`n##########################`n"
}

if ($ForceReleaseUpdate) {
    "`n"
    $removeFilesPath = "/SCREENNETWORK/player/Release_HashSHA256.txt"
    Log -not -msg "`nWlaczony parametr 'ForceReleaseUpdate'`n"
    Log -not -msg "`n##########################`n"
}

$files = ( @(
    if ($UpdateAdmin_W10) {
        $CheckPlayer = $true
        gci "$ftp\update_admin" -Exclude "*update.ps1" | % { [FileObject]@{ File = "$($_.Fullname)"; DestinationPath = "/screennetwork/admin/" } }
    }

    if ($CheckPlayer) {
        [FileObject]@{ File = "$env:USERPROFILE\OneDrive - Screen Network S.A\Dokumenty\REALIZACJA\INSTALATOR\Files\admin\check-player.ps1"; DestinationPath = "/screennetwork/admin/" }
        [FileObject]@{ File = "$env:USERPROFILE\OneDrive - Screen Network S.A\Dokumenty\REALIZACJA\INSTALATOR\Files\admin\functions.psm1"; DestinationPath = "/screennetwork/admin/" }
    }

    if ($ScreenshotService) {
        [FileObject]@{ File = "$ftp\add_ScreenshotService\ScreenshotService.xml"; DestinationPath = "/screennetwork/" }
        [FileObject]@{ File = "$ftp\add_ScreenshotService\ScreenshotService.ps1"; DestinationPath = "/screennetwork/admin/" }
        [FileObject]@{ File = "$ftp\add_ScreenshotService\update.ps1"; DestinationPath = "/screennetwork/" }
    }

    if ($Release) {
        [FileObject]@{ File = "$ftp\Release\v5\Release.zip"; DestinationPath = "/screennetwork/player/" }
    }

    if ($ReleaseWebUpdate) {
        [FileObject]@{ File = "$ftp\ReleaseWebUpdate\update.ps1"; DestinationPath = "/screennetwork/" }
    }

    if ($v3_Release) {
        [FileObject]@{ File = "$ftp\Release\v3\Release.zip"; DestinationPath = "/screennetwork/player/" }
    }

    if ($Config) {
        [FileObject]@{ File = "$ftp\config\$Config\SNPlayer5.config.json"; DestinationPath = "/screennetwork/player/config/" }
    }

    if ($Default) {
        gci "$ftp\default\$Default" | % { [FileObject]@{ File = "$($_.Fullname)"; DestinationPath = "/screennetwork/player/default_content/" } }
    }

    if ($SSH) {
        [FileObject]@{ File = "$ftp\ssh\SSH.zip"; DestinationPath = "/screennetwork/" }
        [FileObject]@{ File = "$ftp\ssh\update.ps1"; DestinationPath = "/screennetwork/" }
    }

    if ($UpdateNetAndRedist) {
        gci "$ftp\update_VCpp_NF461" | % { [FileObject]@{ File = "$($_.Fullname)"; DestinationPath = "/screennetwork/" }}
    }    

    if ($updatePowershell_Win7) {
        [FileObject]@{ File = "$ftp\update_powershell\Win7\update.ps1"; DestinationPath = "/screennetwork/" }
    }

    if ($updatePowershell_Win8) {
        [FileObject]@{ File = "$ftp\update_powershell\Win8\update.ps1"; DestinationPath = "/screennetwork/" }
    }

    if ($CustomFile) {
        [FileObject]@{ File = "$filePath"; DestinationPath = "$sftpPath" }
    }

    if ($OldAndTempFilesDelete) {
        [FileObject]@{ File = "$ftp\DeleteOldFiles\update.ps1"; DestinationPath = "/screennetwork/" }
    }  

    if ($CompleteUpdate) {
        [FileObject]@{ File = "$ftp\CompleteUpdate\update.ps1"; DestinationPath = "/screennetwork/" }
    }  

    if ($CustomAllFilesFromDirectory) {
        gci $directoryPath | % { [FileObject]@{ File = "$($_.Fullname)"; DestinationPath = "$sftpPath" } }
    }
))

if (!(Test-Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll")) {
    "`n"
    Write-Host "Brak zainstalowanego programu WinSCP! Zainstaluj, a nastepnie ponow uruchomienie skryptu" -ForegroundColor red -BackgroundColor Black
    "`n"
    sleep -s 3

    $Browser = new-object -com internetexplorer.application
    $Browser.navigate2("https://winscp.net/eng/downloads.php#additional")
    $Browser.visible = $true
    break
}

if ($null -eq (Get-InstalledModule -name Winscp -ea SilentlyContinue)) {
    Write-Host "Brak zainstalowanego modulu WinSCP, instaluje..." -ForegroundColor Red -BackgroundColor Black    
    sleep -s 1

    $arguments = "Write-host 'Instalacja modulu WinSCP';Install-PackageProvider -Name Nuget -confirm:$false;install-module -name winscp -force"
    Start-Process powershell -Verb runAs -ArgumentList $arguments -Wait
    "`n"
}

try {
    Add-Type -Path ( -Join (${env:ProgramFiles(x86)}, "\WinSCP\WinSCPnet.dll"))
}
catch {
    Write-host $_.Exception.Message -ForegroundColor Red 
    sleep -s 1
    Break   
}

# filter sn
if ($ChooseSN) {
    "`n"
    [string]$snNumbers = Read-Host -Prompt "Wpisz numery SN przedzielajac je przecinkiem (bez spacji)"
    
    if ($snNumbers -eq ($null -or "")) {
        "`n"
        Write-host "Wpisz prawidlowe numery SN!" -ForegroundColor Red
        "`n"
        Break
    }

    if ($snNumbers.StartsWith('sn') -eq $false) {
        $snNumbers = 'sn' + $snNumbers
    }

    $snNumbers = $snNumbers -replace ',', ( -join ([environment]::NewLine, "sn"))
    $filteredServers_Choose = @()

    foreach ($value in $servers) {
        if ($snNumbers.Contains($value.name)) {
            $filteredServers_Choose = $filteredServers_Choose + $value
        }
    }

    $servers = $filteredServers_Choose
}

# Przedzial numerow SN
if ($CollectionSN) {
    "`n"
    [int]$snNumbersStart = Read-Host -Prompt "Wpisz numer poczatkowy zbioru (bez 'sn')"
    [int]$snNumbersEnd = Read-Host -Prompt "Wpisz numer koncowy zbioru (bez 'sn')"

    if (($snNumbersEnd -lt $snNumbersStart) -or ($snNumbersStart -eq $snNumbersEnd) -or (($snNumbersStart -or $snNumbersEnd) -eq ($null -or 0))) {
        "`n"
        Write-Host "Wprowadz prawidlowe wartosci, poczatkowa liczba nie moze byc mniejsza ani rowna koncowej." -ForegroundColor Red
        "`n"
        Break
    }

    [array]$filteredServers_Collection = $snNumbersStart..$snNumbersEnd 
    
    for ($i = 0; $i -le $filteredServers_Collection.Count - 1; $i ++) {
        $filteredServers_Collection[$i] = $filteredServers_Collection[$i].ToString()
        
        if ($filteredServers_Collection[$i].Length -eq 1) {
            $filteredServers_Collection[$i] = "00" + $filteredServers_Collection[$i]
        }

        if ($filteredServers_Collection[$i].Length -eq 2) {
            $filteredServers_Collection[$i] = "0" + $filteredServers_Collection[$i]
        }

        $filteredServers_Collection[$i] = "sn" + $filteredServers_Collection[$i]
    }

    foreach ($value in $servers) {
        if ($filteredServers_Collection.Contains($value.name)) {
            $filteredServers_Collection_Values = [array]$filteredServers_Collection_Values + $value
        }
    }

    $servers = $filteredServers_Collection_Values
}

# Pomin SN
if ($ExcludeSN) {
    "`n"
    [string]$snNumbers = Read-Host -Prompt "Wpisz numery SN, ktore maja zostac pominiete, przedzielajac je przecinkiem (bez spacji)"

    if ($snNumbers.StartsWith('sn') -eq $false) {
        $snNumbers = 'sn' + $snNumbers
    }

    $snNumbers = $snNumbers -replace ',', ( -join ([environment]::NewLine, "sn"))
    $filteredServers_Exclude = @()

    $servers | % { 
        if ($snNumbers -notmatch $_.name) { 
            $filteredServers_Exclude = [array]$filteredServers_Exclude + $_ 
        } 

    }

    $servers = $filteredServers_Exclude
}

# Podziel SN
if (($SplitSN -ne 0) -and ($SplitSN -lt $servers.Count)) {
    ""
    Write-Host "Dziele na listy o maksymalnej ilosci rekordow: $SplitSN"
    ""
    $MyInvocation.BoundParameters.keys | ForEach {
        if ( $skipParameters -notcontains $_ ) {
            $boundParameters += $_
        }
    }
    
    ($MyInvocation.MyCommand.Parameters ).Keys | ForEach {
        if (( $boundParameters -notcontains $_ ) -and ( $builtinParameters -notcontains $_ ) -and ( $skipParameters -notcontains $_ )) {
            $val = (Get-Variable -Name $_ -EA SilentlyContinue).Value
            if ((( $val -ne $false ) -and ($_ -ne "Config")) -and (( $val -ne $false ) -and ($_ -ne "Default"))) {
                $boundParameters += " -$_"
            }
            elseif ((( $val -ne $false ) -and ($_ -eq "Config")) -or (( $val -ne $false ) -and ($_ -eq "Default"))) {
                $boundParameters += " -$_ $val"
            }
        }
    }

    $boundParameters = $boundParameters | % {-join(" -", $_)} | Out-String
    $boundParameters = $boundParameters.replace("`r`n", "" )

    
    $xmlPath = split-path $myinvocation.mycommand.definition -Parent
    $xmlCounter = 0
    $outArray = @()

    foreach ($serv in $servers) {
        $outArray = [array]$outArray + $serv

        if ($outArray.count -eq $SplitSN) {
            $xmlFullname = "$xmlPath/servers_$xmlCounter.xml"
            $outArray | Export-Clixml $xmlFullname -Force
            $xmlCounter++
            $outArray = @()
        }
    }
   
    if ($outArray.count -ne 0) {
        $xmlFullname = "$xmlPath/servers_$xmlCounter.xml"
        $outArray | Export-Clixml $xmlFullname
        $outArray = @()
    }

    gci $xmlPath -Filter *.xml | % {
        $arg = ( -join ("& ", $myinvocation.mycommand.definition), $boundParameters, "-InputServers `"", $_.FullName, "`"")
        Start-Process powershell -ArgumentList $arg
    }

    Remove-Item $global:log -Force
    Exit
}

$files | % { 
    if ($_.file -eq "") {
        Write-Warning "`nSciezka do pliku jest pusta, pomijanie...`n"
    }
    elseif ((test-path $_.file) -eq $true) { 
        $files2 = [array]$files2 + $_ 
    } 
    else { 
        Write-Warning "`nBledna sciezka do pliku  $($_.File),  usuwanie...`n"
    } 
}

$files = $files2

if ($files.count -gt 0) {
    $filesInArray = $true

    Log -not -msg "`nWysylane pliki:`n"
    for ($i = 0; $i -lt $files.count; $i++) {
        Log -not -msg ( -join ("$($i + 1). ", $files[$i].File, "   wysylane do:   ", $files[$i].DestinationPath))
    }

    Log -not -msg "`n##########################`n"
}
else {
    $filesInArray = $false
}

"`n"
foreach ($server in $servers) {
    $script:lastFileName = $Null
    $sn = $server.name
    $ip = $server.ip
    Write-host " "$sn", "$ip"  " -ForegroundColor Green -BackgroundColor Black
    Log -not -msg "$sn - $ip"

    if ($ip -ne "NULL") {
        try { 
            $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
                Protocol   = [WinSCP.Protocol]::ftp
                HostName   = $ip
                PortNumber = 0
                UserName   = "***"
                Password   = "***"
            }
 
            $session = New-Object WinSCP.Session
 
            try {
                $transferResult = $null
                $result = @()
                $sendedFiles = @()
                $session.add_FileTransferProgress( { FileTransferProgress($_) } )

                $session.Open($sessionOptions)
 
                if ($DeleteFiles -or $ForceReleaseUpdate) {
                    $deleteResults = $session.RemoveFiles($removeFilesPath)

                    if ($deleteResults.Removals.count -ne 0){
                        $deleteResults.Removals | % {
                            $fnLeaf = Split-Path ($_.FileName) -Leaf -ErrorAction SilentlyContinue

                            if ($_.Error -eq $null) {
                                Write-Host (-join($(Get-Date), " [Remove]: ", $fnLeaf))
                                Log -del -msg $fnLeaf
                            }
                            elseif ($_.Error -ne $null) {
                                Write-Host (-join($(Get-Date), " [Remove]: ", $fnLeaf, ", ERROR: ", $_.Error))
                                Log -e -msg (-join($fnLeaf, " - ", $_.Error))
    
                                $failedRemoving = [array]$failedRemoving + (New-Object psobject -Property ([ordered]@{Komputer = $sn ; Plik = $fnLeaf}))
                            }
                        }
    
                        if ($deleteResults.IsSuccess -eq $true) {
                            $done = [array]$done + (New-Object psobject -Property ([ordered]@{Komputer = $sn }))
                        }
                    }
                    elseif (($deleteResults.Removals.count -eq 0) -and $deleteResults.IsSuccess -eq $false) {
                        Write-Host (-join($(Get-Date), " [Remove]: Plik lub sciezka nie zostala odnaleziona"))
                        Log -e -msg "Plik lub sciezka nie zostala odnaleziona"
                    }
                }

                if ($Restart) { $session.CreateDirectory("/restart.lock") }
                
                if ($Restart_Player) { $session.CreateDirectory("/restart_player.lock") }

                if ($filesInArray) {
                    $transferOptions = New-Object WinSCP.TransferOptions
                    $transferOptions.OverwriteMode = [WinSCP.OverwriteMode]::Overwrite

                    foreach ($fileToSend in $files) {
                        $sendedFiles = [array]$sendedFiles + $session.PutFiles($fileToSend.File, $fileToSend.DestinationPath, $False, $transferOptions)
                    }

                    $sendedFiles | % {
                        $f = Split-Path (Convert-Path -path $_.Transfers.Filename) -Leaf
                        $hash = [ordered]@{File = $f; TransferSuccess = $_.IsSuccess }
                        $result = [array]$result + (New-Object psobject -Property $hash)
                    }

                    if ($null -ne $transferResult) {
                        $transferResult.Check() 
                    }

                    "`n"
                    ($result | Out-String).Trim() | ft -AutoSize
                
                    foreach ($r in $result) { 
                        if ($r.TransferSuccess -eq $False) {
                            Log -e -msg $r.File
                            Write-Host "Przesylanie $($r.File) nie powiodlo sie!" -ForegroundColor Red -BackgroundColor Black
                            $failed = [array]$failed + (New-Object psobject -Property ([ordered]@{Komputer = $sn; Plik = $r.File }))
                        }
                        else { Log -msg $r.File }
                    }

                    if (($result.TransferSuccess | select -Unique).Count -eq 1) {
                        $done = [array]$done + (New-Object psobject -Property ([ordered]@{Komputer = $sn }))
                    }
                }  
            } 
            catch {
                $eMsg = (-join("`n", $_.Exception.Message, "`n`nLine ", $error[0].InvocationInfo.ScriptLineNumber, " : " + ($error[0].InvocationInfo.Line | Out-String).Trim() ))
                Log -e -msg $eMsg
                Write-Host "Wystapil blad" -ForegroundColor Red -BackgroundColor Black
                Write-Host $eMsg
                $fResend = @()
                $fNames = @()

                foreach ($file in $files) {
                    $fNames = [array]$fNames + (split-path $file.file -leaf)
                    $fResend = [array]$fResend +  (New-Object psobject -Property ([ordered]@{File = $file.file; DestinationPath = $file.DestinationPath}))
                }

                $failed = [array]$failed + (New-Object psobject -Property ([ordered]@{Komputer = $sn; IPv4 = $ip; Plik = $fNames; FilsToResend = $fResend}))
            } 
            finally {
                if ($Null -ne $script:lastFileName) {
                    Write-Host
                }

                $session.Dispose()
            }
            
            "`n"
        }
        catch {
            Write-Host "`nError: $($_.Exception.Message)"
            Log -e -msg "Connection error: $($_.Exception.Message)"
        }
    }
    else {
        $notConnected = [array]$notConnected + (New-Object psobject -Property ([ordered]@{Komputer = $sn}))
        Write-Host "Komputer nie jest polaczony z VPN" -ForegroundColor Red -BackgroundColor Black
        Log -e -msg "Komputer nie jest polaczony z VPN"
        "`n"
    }

    Log -not -msg ""
}

Log -not -msg "##########################`n"

if ($done.Count -ne 0) {
    write-host "`n`nDONE:`n" -ForegroundColor Green -BackgroundColor Black
    Log -not -msg "DONE:"
    foreach ($d in $done) {
        if (!($allDone -like "*$($d.Komputer)*")) {
            $allDone += (-join($d.Komputer, [System.Environment]::NewLine))
        }
    }

    Write-host $allDone.Trim()
    Log -not -msg $allDone.Trim()

    write-host ""
    Log -not -msg ""
}

if ($notConnected.Count -ne 0) {
    write-host "`n`nNiepolaczone komputery`n" -ForegroundColor Red -BackgroundColor Black
    Log -not -msg "Niepolaczone komputery:"
    $notConnected | % {
        write-host $_.Komputer
        Log -not -msg $_.Komputer
    }

    write-host ""
    Log -not -msg ""
}

if ($failed.Count -ne 0) {
    write-host "`n`nNiepowodzenie wysylania plikow`n" -ForegroundColor Red -BackgroundColor Black
    Log -not -msg "Niepowodzenie wysylania plikow:"
    $failed | % {
        $f = (($_.Plik | % { "$_ "}) | Out-String)
        Write-Host (-join($_.Komputer, " - ", $f))
        Log -not -msg $a
    }

    write-host ""
    Log -not -msg ""

    foreach ($p in $failed) {
        if (!($allFailedString -like "*$($p.Komputer)*")) {
            $allFailedString += (-join($p.Komputer, ","))
        }
    }

    $allFailedString.TrimEnd(',').Trim()
}

if ($failedRemoving.Count -ne 0) {
    write-host "`n`nNiepowodzenie usuwania plikow`n" -ForegroundColor Red -BackgroundColor Black
    Log -not -msg "Niepowodzenie usuwania plikow:"
    $failedRemoving | % {
        $f = (($_.Plik | % { "$_ "}) | Out-String)
        Write-Host (-join($_.Komputer, " - ", $f))
        Log -not -msg $a
    }

    write-host ""
    Log -not -msg ""
}

if ($InputServers -ne $null) {
    Pause
}