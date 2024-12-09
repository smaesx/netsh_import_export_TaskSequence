##########################################
# SET VARIABLES FOR TS ENVIRONMENT
##########################################
$TSEnv = New-Object -COMObject "Microsoft.SMS.TSEnvironment"

$logPath = $TSEnv.Value('OSD_LogPath')
$logFolder = $TSEnv.Value('OSD_LogFolderName')
$logDirectory = "$logPath$logFolder"

##########################################
# EXPORT ETHERNET ADAPTER CONFIGURATION
##########################################
Start-Transcript "$logDirectory\Netsh_Export.log"
try {
    # START SERVICE DOT3SVC
    if((Get-Service -Name "dot3svc").Status -ne "Running") {
        Write-Output "Starting service dot3svc..."
        Start-Service -Name dot3svc -ErrorAction Stop
    }
    else {
        Write-Output "Service dot3svc is already running."
    }

    # GET NIC NAME
    $NICName = Get-NetAdapter -ErrorAction Stop | Where-Object {$_.Status -eq "Up" -and $_.InterfaceAlias -like "*Ethernet*"} | Select-Object -ExpandProperty Name
    Write-Output "NICName: $NICName"

    if($null -ne $NICName -and $NICName -ne "") {
        # EXPORT NETWORK CONFIGURATION
        $Args = @{
            FilePath = "netsh"
            ArgumentList = "lan export profile folder=`"$logDirectory`" interface=`"$NICName`""
            Wait = $true
            NoNewWindow = $true
            ErrorAction = "Stop"
        }
        Write-Output "Exporting network configuration..."
        Start-Process @Args
        Write-Output "Network configuration exported to $logDirectory\$NICName.xml"
    }
}
catch {
    Write-Error $_
}
finally {
    Stop-Transcript
}
