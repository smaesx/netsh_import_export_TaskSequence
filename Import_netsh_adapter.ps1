##########################################
# SET VARIABLES FOR TS ENVIRONMENT
##########################################
$TSEnv = New-Object -COMObject "Microsoft.SMS.TSEnvironment"

$logPath = $TSEnv.Value('OSD_LogPath')
$logFolder = $TSEnv.Value('OSD_LogFolderName')
$logDirectory = "$logPath$logFolder"

##########################################
# IMPORT ETHERNET ADAPTER CONFIGURATION
##########################################
Start-Transcript "$logDirectory\Netsh_Import.log"
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
        
        if(-not (Test-Path -Path "$logDirectory\$NICName.xml")) {
            Write-Output "Network configuration file $logDirectory\$NICName.xml not found. Skipping import."
            break
        }
        
        # IMPORT NETWORK CONFIGURATION
        $Args = @{
            FilePath = "netsh"
            ArgumentList = "lan add profile filename=`"$logDirectory\$NICName.xml`" interface=`"$NICName`""
            Wait = $true
            NoNewWindow = $true
            ErrorAction = "Stop"
        }
        Write-Output "Importing network configuration..."
        Start-Process @Args
        Write-Output "Network configuration imported from $logDirectory\$NICName.xml"

        # DISABLE NETWORK ADAPTER
        Write-Output "Disabling and enabling network adapter $NICName..."
        Disable-NetAdapter -Name $NICName -Confirm:$False -ErrorAction Stop
        Write-Output "Network adapter $NICName disabled."

        # ENABLE NETWORK ADAPTER
        Write-Output "Enabling network adapter $NICName..."
        Enable-NetAdapter -Name $NICName -ErrorAction Stop
        Write-Output "Network adapter $NICName enabled."
    }
}
catch {
    Write-Error $_
}
finally {
    Stop-Transcript
}
