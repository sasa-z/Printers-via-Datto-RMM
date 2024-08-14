
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted 
Set-ExecutionPolicy -ExecutionPolicy Bypass -Force

function ImportFunction-FromGit
{
    <#
    .SYNOPSIS
        This script is intended to import scripts from GitHub.
    .DESCRIPTION
        This script is intended to import scripts from GitHub as function into PowerShell
    .PARAMETER Url
        The parameter Url specifies the link to the raw GitHub content.
    .PARAMETER FunctionName
        The parameter FunctionName can be used to specify the name of the function.
    .PARAMETER AlreadyFunction
        The parameter AlreadyFunction shall be used, when the content contains functions.
    .EXAMPLE
        # import Get-Autodiscover function from PowerShell script
        ImportFunction-FromGit -Url 'https://raw.githubusercontent.com/IngoGege/Get-Autodiscover/main/Get-Autodiscover.ps1'
        # import Get-AccessToken function from PowerShell script with name Get-AADToken
        ImportFunction-FromGit -Url 'https://raw.githubusercontent.com/IngoGege/Get-AccessToken/master/Get-AccessToken.ps1' -FunctionName Get-AADToken
    .NOTES

    .LINK
        https://ingogegenwarth.wordpress.com/
    #>
    [CmdletBinding()]
    param(
        [parameter(
            mandatory = $true,
            Position = 0)]
        [System.Uri]
        $Url,

        [parameter(
            mandatory = $false,
            Position = 1)]
        [System.String]
        $FunctionName,

        [parameter(
            mandatory = $false,
            Position = 2)]
        [System.Management.Automation.SwitchParameter]
        $AlreadyFunction

    )

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        # retrieve code
        $code = Invoke-RestMethod -Method GET -Uri $Url

        if ([System.String]::IsNullOrEmpty($FunctionName))
        {
            Write-Verbose "FunctionName not given..."
            $FunctionName = ($Url.AbsoluteUri.ToString().Split('/')[ -1]).Split('.')[0]
            Write-Verbose "Using:$($FunctionName)..."
        }

        if ($AlreadyFunction)
        {
            # create temporary file for import
            $fileName = "$(Get-Random).psm1"
            $tempFile = New-Item -Path $env:TEMP -Name $fileName -Value $code -Force 
            Import-Module $tempFile.FullName -Global -DisableNameChecking
            # cleanup of temporary file
            Remove-Item $tempFile -Force
        }
        else
        {
            Invoke-Expression "function global:$($FunctionName) { $($code) }"
        }
    }
    catch {
        $_
    }
    
}


ImportFunction-FromGit -Url "https://github.com/sasa-z/DattoRMMPSHelper/raw/main/DattoRMMHelper.psm1" -AlreadyFunction

$ifUserLoggedInCheck  = (Get-WmiObject -ClassName Win32_ComputerSystem).Username


if(-not $ifUserLoggedInCheck){
    send-Log -logText "_Warning. User is not logged in and we can't set printer as default or set printer configuration" -addDashes Below -type Warning -addTeamsMessage

}

$Global:ScriptName ="Brother_MFC-L9670CDN"
$Global:ToastNotificationAppLogo = 'Printer.png'
$Global:ToastNotificationHeader = "Brother MFC-L9670CDN"
$ScriptFolderLocation = "$env:rootScriptFolder\$scriptName"
$PrintingDefaults = $env:PrintingDefaults 

$global:EnvDattoVariablesValuesHashTable = @{}
#add/change Datto variables name and descriptions below
$EnvDattoVariablesValuesHashTable.Add("$($env:Action)", "What action you want to do?") #change this variable value according to Datto variables in this case, replace Action, and description etc..
$EnvDattoVariablesValuesHashTable.Add("$($env:PrinterIP)", "Add printer IP address") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("$($env:PrinterName1)", "Add Printer Name 1") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("$($env:PrintingDefaults1)", "Select Printing Defaults 1") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("$($env:PrinterName2)", "Add Printer Name 2") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("$($env:PrintingDefaults2)", "Select Printing Defaults 2") #change this according to Datto variables
$EnvDattoVariablesValuesHashTable.Add("$($env:SetPrinterAsDefault)", "Set printer as default?") #change this according to Datto variables


add-ScriptWorkingFoldersAndFiles
get-runAsUserModule
get-BurntToastModule
remove-oldToastNotifications

# printer driver if we need to download $URL = "https://download.brother.com/welcome/dlf105599/Y20E_C1-hostm-C1.EXE"

$inf = "BRPRC19A.INF"
$InfLocation = "$ScriptFolderLocation\Driver\$inf" #this is driver we add into drivers store
$driverName = "Brother MFC-L9610CDN series" #can be found in inf file under [driver name] section and it should be exactly the same here
$PrinterPortIP = "$env:PrinterIP" # IP address
$printerName1 = "$env:PrinterName1"
$printerName2 = "$env:PrinterName2"
$SetPrinterAsDefault = $env:SetPrinterAsDefault
# $defaultPrinterDisplayName1 =  "Brother MFC-L9670CDN - B&W" #is nothing is set in Datto, we use this one
# $defaultPrinterDisplayName2 =  "Brother MFC-L9670CDN - COLOR" #is nothing is set in Datto, we use this one

#region if printers exist
if ($env:PrinterName1){
    try{$printer1Exists = get-printer -Name $printerName1 -ErrorAction stop}catch{}
}
if ($env:PrinterName2){
    try{$printer2Exists = get-printer -Name $printerName2 -ErrorAction stop}catch{}
}

#endregion if printers exist

#region OS
$OS = (Get-WmiObject Win32_OperatingSystem).Caption

$ProgressPreference = 'SilentlyContinue'

if($OS -like "*Windows 10*"){
    $DriverURL = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Win10/Driver.zip"
    send-log -logText "Driver for Windows 10 machine"

}elseif($OS -like "*Windows 11*"){
    $DriverURL = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Win11/Driver.zip"
    send-log -logText "Driver for Windows 11 machine"
}else{
    send-log -logText "_Warning. OS not supported (Windows 10 or 11 only)"
    exit 
}
#endregion OS

#region configuration and port names
<#
we get this file below by installing driver on test machine, set printing default and export default printing configuration 
by (Get-PrintConfiguration).PrintTicketXML | out-file "Mono-1-Sided.xml" and upload to datto script
#>

if ($env:PrinterName1){

    if ($env:PrintingDefaults1 -eq "BW1Sided"){
        $ConfigURL1 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/BW1Sided.xml"
        $portName1 =  $PrinterPortIP + "_BW"
        
        send-Log -logText "portName1: $portName1"
    }elseif($env:PrintingDefaults1 -eq "Auto1Sided"){
        $ConfigURL1 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/Auto1Sided.xml"
        $portName1 =  $PrinterPortIP + "_AUTO"
        
        send-log "portName1: $portName1"
    }elseif($env:PrintingDefaults1 -eq "Color1Sided"){
        $ConfigURL1 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/Color1Sided.xml"
        $portName1 =  $PrinterPortIP + "_COLOR"
        
        send-log "portName1: $portName1"
    
    }else{
        $portName1 =  $PrinterPortIP + "_AUTO" 
        send-log "portName1: $portName1"
    }

}

if ($env:PrinterName2){ 
    if ($env:PrintingDefaults2 -eq "BW1Sided"){
        $ConfigURL2 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/BW1Sided.xml"
        $portName2 =  $PrinterPortIP + "_BW"
        
        send-log -logText "portName2: $portName2"
    }elseif($env:PrintingDefaults2 -eq "Auto1Sided"){
        $ConfigURL2 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/Auto1Sided.xml"
        $portName2 =  $PrinterPortIP + "_AUTO"
        
        send-log -logText "portName2: $portName2"
    
    }elseif($env:PrintingDefaults2 -eq "Color1Sided"){
        $ConfigURL2 = "https://raw.githubusercontent.com/sasa-z/Printers/main/Brother%20MFC-L9670CDN/Configurations/Color1Sided.xml"
        $portName2 =  $PrinterPortIP + "_COLOR"
        
        send-log -logText "portName2: $portName2"
    }else{
        $portName2 =  $PrinterPortIP + "_AUTO"
        send-log -logText "portName2: $portName2"
    }
    

}


#endregion configuration and port names


#region copy driver and config to local machine
try{

    Invoke-WebRequest $DriverURL -OutFile "$ScriptFolderLocation\Driver.zip" 
    if($ConfigURL1){Invoke-WebRequest $ConfigURL1 -OutFile "$ScriptFolderLocation\PrinterConfiguration1.xml"  }
    if($ConfigURL2){Invoke-WebRequest $ConfigURL2 -OutFile "$ScriptFolderLocation\PrinterConfiguration2.xml" }
}catch{
    send-Log -logText "Failed to copy Driver.zip or config file to local machine" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message
    send-CustomToastNofication -text "Failed to copy files" -type Error
    send-CustomFinalToastNotification -SendToTeams
    exit 1
}
#endregion copy driver to local machine


#region expand driver
try{
    Expand-Archive -Path "$ScriptFolderLocation\Driver.zip" -DestinationPath "$ScriptFolderLocation\Driver" -Force -ErrorAction Stop
}catch{
    send-Log -logText "Failed to expand Driver.zip to local machine" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message
    send-CustomToastNofication -text "Failed to expand Driver.zip" -type Error
    send-CustomFinalToastNotification -SendToTeams
    exit 1
}
#endregion expand driver

#region IP address validation

if ($env:action -eq 'add'){

    $pattern = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"

    if($PrinterPortIP -notmatch $pattern){
    send-log -logText "_Warning. Printer IP address is not valid" -type Warning
    exit
    }
}elseif($env:action -eq 'update'){
    if($PrinterPortIP -notmatch $pattern -and $PrinterPortIP){
    send-log -logText "_Warning. Printer IP address is not valid" -type Warning
    exit
    }
}

#endregion IP address validation

#region checks for action ADD
if ($env:action -eq 'add'){

    #region if both printer names are empty, exit
        if(-not $env:PrinterName1 -and -not $env:PrinterName2){
            send-log -logText "_Warning. Both printer names are empty" -type Warning
            exit
        }

        if($env:PrinterName1 -eq $env:PrinterName2){
            send-log -logText "_Warning. Both printers have the same name" -type Warning
            exit
        }
    #endregion if both printer names are empty, exit

    #region if both printers exist and action is add, exit
    If ($printerName1){
        $Printer1Exists = get-printer -Name $printerName1
     }

     If ($printerName2){
        $Printer2Exists = get-printer -Name $printerName2
     }

     if ($Printer1Exists -and $Printer2Exists){
        send-log -logText "_Warning. Printer $printerName1 and $printerName2 already exist and you selected action: ADD" -type Warning
        exit
     }

    #endregion if both printers exist and action is add, exit

    #region IP is missing
    if (-not $PrinterPortIP){
        send-log -logText "_Warning. Printer IP address is missing" -type Warning
        exit
    }
    #endregion IP is missing


    }

#endregion checks for action ADD

#region checks for action UPDATE
if ($env:action -eq 'update'){

    #region if both printer names are empty, exit
        if(-not $env:PrinterName1 -and -not $env:PrinterName2){
            send-log -logText "_Warning. Both printer names are empty" -type Warning
            exit
        }

        if($env:PrinterName1 -eq $env:PrinterName2){
            send-log -logText "_Warning. Both printers have the same name" -type Warning
            exit
        }


    }

#endregion checks for action UPDATE









#endregion checks


#region adding driver to driver store
send-CustomToastNofication -text "Installing print driver"

try{
    start-process pnputil -ArgumentList "/add-driver", $InfLocation, "/install" -wait -ErrorAction Stop
    send-Log -logText "Added driver into driver store"

}catch{
    send-Log -logText "Failed to add driver into driver store" -type Error -addTeamsMessage
    send-CustomToastNofication -text "Failed to add driver" -type Error
    send-CustomFinalToastNotification -SendToTeams
    exit 1

}
#endregion adding driver to driver store

#region add printer driver
try{
    $CheckDriver = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
    if($CheckDriver){
        send-Log -logText "Printer driver $($driverName) already exists"
        
    }else{
        Add-PrinterDriver -Name $driverName -ErrorAction Stop
        send-Log -logText "Added printer driver successfully"
    }

}catch{
    send-Log -logText "Failed to add or check printer driver" -type Error -addTeamsMessage
    send-CustomToastNofication -text "Failed to add driver" -type Error
    send-CustomFinalToastNotification -SendToTeams
    exit 1
}

#endregion printer driver

#region Printer port

if ($env:action -eq 'update' -and -not $env:PrinterIP){
    send-log -logText "Printer port not set and action is UPDATE. Skipping adding printer port"
}else{

    if ($printerName1){

        try{$CheckPrinterPort = Get-PrinterPort -Name $portName1 -ErrorAction SilentlyContinue}catch{}
    
        if($CheckPrinterPort -and ($CheckPrinterPort.PrinterHostAddress -eq $PrinterPortIP)){ #printer port exist with same IP
            send-Log -logText "Printer port $($portName1) already exists" -addDashes above
    
        }elseif($CheckPrinterPort){ #printer port exist but with different IP
            send-Log -logText "Printer port $($portName1) with IP: $($PrinterPortIP) already exists but with different IP" -type Warning -addTeamsMessage -addDashes above
            send-CustomToastNofication -text "Port exists with different IP" -type Warning
            send-CustomFinalToastNotification -SendToTeams
            exit 
    
        }else{ #printer port does not exist, adding it
    
            try{
                
                Add-PrinterPort -Name "$portname1" -PrinterHostAddress $PrinterPortIP -ErrorAction Stop
                send-Log -logText "Added printer port $($portName1) with IP: $($PrinterPortIP) successfully"
            }catch{
                send-Log -logText "Failed to add printer port $($PrinterPortIP)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message -addDashes above
                send-CustomToastNofication -text "Failed to add printer port" -type Error
                send-CustomFinalToastNotification -SendToTeams
                exit 1
            }
        }
    
    }
        
    
        if ($env:PrinterName2){
    
            
            try{$CheckPrinterPort = Get-PrinterPort -Name $portName2 -ErrorAction SilentlyContinue}catch{}
    
            if($CheckPrinterPort -and ($CheckPrinterPort.PrinterHostAddress -eq $PrinterPortIP)){ #printer port exist with same IP
                send-Log -logText "Printer port $($portName2) already exists" -addDashes above
        
            }elseif($CheckPrinterPort){ #printer port exist but with different IP
                send-Log -logText "Printer port $($portName2) with IP: $($PrinterPortIP) already exists but with different IP" -type Warning -addTeamsMessage
                send-CustomToastNofication -text "Port exists with different IP" -type Warning
                send-CustomFinalToastNotification -SendToTeams
                
        
            }else{ #printer port does not exist, adding it 
        
                try{
                    
                    Add-PrinterPort -Name "$portname2" -PrinterHostAddress $PrinterPortIP -ErrorAction Stop
                    send-Log -logText "Added printer port $($portName2) with IP: $($PrinterPortIP) successfully"
                }catch{
                    send-Log -logText "Failed to add printer port $($PrinterPortIP)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message
                    send-CustomToastNofication -text "Failed to add printer port" -type Error
                    send-CustomFinalToastNotification -SendToTeams
                    exit 1
                }
            }
        }


}



#endregion Printer port

#region add or update printer

if ($env:action -eq "add" -and $env:PrinterName1){

    if($printer1Exists){       
        send-Log -logText "Printer $($printerName1) already exists and can't be added"  -addTeamsMessage -addDashes above
        send-CustomToastNofication -text "$($printerName1) already exists"

    }else{ #printer doesn't exist
        send-Log -logText "Printer $($printerName1) does not exist. Adding printer" -addDashes above

        try{
            send-Log -logText "Adding printer $($printerName1), with driver $($driverName) and port $($portName1)" -addDashes Below
            Add-Printer -DriverName $driverName -Name $printerName1 -PortName $portName1 -ErrorAction Stop
            send-Log -logText "$($printerName1) added successfully" -addteamsMessage -addDashes below
            send-CustomToastNofication -text "$($printerName1) added"
    
        }catch{
            send-Log -logText "Failed to add printer $($printerName)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message -addDashes above
            send-CustomToastNofication -text "Failed to add $($printerName)" -type Error
            send-CustomFinalToastNotification -SendToTeams
            exit 1
        }

    }

}elseif($env:action -eq "update" -and $env:PrinterName1){ 
    
    if($printer1Exists){       

        if ($PrinterPortIP){ #port specified and we need to update it
            send-Log -logText "Updating printer port $($portName1) with IP: $($PrinterPortIP)"
            try{
                Set-Printer -Name $printerName1 -DriverName $driverName -PortName $portName1 -ErrorAction Stop
    
            }catch{
                send-Log -logText "Failed to update printer $($printerName1) with port $($portName1)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message -addDashes above
                send-CustomToastNofication -text "Failed to update printer port" -type Error
                send-CustomFinalToastNotification -SendToTeams
                exit 1
            }
            
    
        }

    }else{
        send-Log -logText "Printer $($printerName1) does not exist and we can't update printer port" -addDashes above
    }
  

}



if ($env:action -eq "add" -and $env:PrinterName2){

        if($printer2Exists ){       
            send-Log -logText "Printer $($printerName2) already exists and can't be added"  -addTeamsMessage -addDashes above
            send-CustomToastNofication -text "$($printerName2) already exists"
       
            
        }else{ #printer doesn't exist
            send-Log -logText "Printer $($printerName2) does not exist. Adding printer" -addDashes above

            try{
                send-Log -logText "Adding printer $($printerName2), with driver $($driverName) and port $($portName2)"
                Add-Printer -DriverName $driverName -Name $printerName2 -PortName $portName2 -ErrorAction Stop
                send-Log -logText "Added printer $($printerName2) successfully" -addteamsMessage -addDashes below
                send-CustomToastNofication -text "$($printerName2) added"
        
            }catch{
                send-Log -logText "Failed to add printer $($printerName2)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message
                send-CustomToastNofication -text "Failed to add $($printerName2)" -type Error
                send-CustomFinalToastNotification -SendToTeams
                exit 1
            }

        }


}elseif($env:action -eq "update" -and $env:PrinterName2){ 
    
    if($printer2Exists){       

        if ($PrinterPortIP){ #port specified and we need to update it
            send-Log -logText "Updating printer port $($portName2) with IP: $($PrinterPortIP)"
            try{
                Set-Printer -Name $printerName2 -DriverName $driverName  -PortName $portName2 -ErrorAction Stop
    
            }catch{
                send-Log -logText "Failed to update printer $($printerName2) with port $($portName2)" -type Error -addTeamsMessage -catchMessage $error[0].Exception.Message -addDashes above
                send-CustomToastNofication -text "Failed to update printer port" -type Error
                send-CustomFinalToastNotification -SendToTeams
                exit 1
            }
            
    
        }

    }else{
        send-Log -logText "Printer $($printerName2) does not exist and we can't update printer port" -addDashes above
    }
  

}





    
#endregion add or update printer

#region set printer as default

if ($env:SetPrinterAsDefault -notlike "skip"){

    if ($SetPrinterAsDefault -like "Printer1" -or $SetPrinterAsDefault -like "Printer2"){

        if ($SetPrinterAsDefault -like "Printer1"){

            $printerName = $printerName1
            if ($env:PrinterName1 -and $printer1Exists -and $env:action -eq 'add'){ #don't set printer as default if printer exists and action is add
                $proceed = $false
            }elseif(-not $env:PrinterName1){
                $proceed = $false
            }else{
                $proceed = $true

            }
        
        }elseif($SetPrinterAsDefault -like "Printer2"){
            $printerName = $printerName2
            if ($env:PrinterName2 -and $printer2Exists  -and $env:action -eq 'add'){ #don't set printer as default if printer exists and action is add
                $proceed = $false
            }elseif(-not $env:PrinterName2){
                $proceed = $false
            }else{
                $proceed = $true
            }
        }
        
    if ($proceed){

        send-Log -logText "Setting $($printerName) as default printer"

        #we can't read parent variables with invoke-ascurrentuser so we need to place them into file
        try{remove-item "DefaultPrinter.txt" -Path "c:\yw-data\automate" -force  -ErrorAction SilentlyContinue }catch{}
        New-Item -Path c:\yw-data\automate -Name "DefaultPrinter.txt" -ItemType "file" -Value $printerName  | out-null
        (Get-Item c:\yw-data\automate\DefaultPrinter.txt).Attributes = "Hidden"
   
        Invoke-AsCurrentUser -ScriptBlock {
       
           $PrinterName = Get-Content "c:\yw-data\automate\DefaultPrinter.txt"
           $printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$PrinterName'"
           Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter
           Remove-Item "c:\yw-data\automate\DefaultPrinter.txt" -force
        }
        send-Log -logText "Printer $($printerName) set as default" -addDashes Below 
        send-CustomToastNofication -text "Printer $($printerName) set as default"

    }else{
        send-log -logText "Printer $($printerName) exists and we can't set it as default. Use action UPDATE" -addDashes below
    }

   
    }elseif($SetPrinterAsDefault -eq "WindowsManage"){
        send-Log -logText "Letting Widnows to manage default printers"

        Invoke-AsCurrentUser -ScriptBlock {
       
           Set-ItemProperty -path "HKCU:\\Software\Microsoft\Windows NT\CurrentVersion\Windows" -name "LegacyDefaultPrinterMode" -value "00000000"
         }

    }else{
        
        send-Log -logText "Skipping setting printer as default"
    }


    
}
    



   

#endregion set printer as default


#region printer configuration

if ($env:PrinterName1 -and ($env:printingDefaults1 -notlike 'skip')){ #skip it action is skip or printer name is not set

    if ($env:action -eq 'add' -and $printer1Exists){ #skip if printer already exists with action ADD
        send-Log -logText "Skipping printer configuration for $($printerName1) as it already exists" -addDashes below
    
    }elseif(($env:action -eq 'update' -and $printer1Exists) -or ($env:action -eq 'add' -and -not $printer1Exists)){ #proceed if action is UPDATE and printer exists or ADD and printer does not exist


        send-Log -logText "Restoring printer configuration for $($printerName1): $($env:printingDefaults1)" -addDashes Below
    
        $config = Get-Content "$ScriptFolderLocation\PrinterConfiguration1.xml" | out-string
    
        if ($config){
    
            #we can't read parent variables with invoke-ascurrentuser so we need to place them into file
            try{remove-item "PrinterName1.txt" -Path "c:\yw-data\automate" -force  -ErrorAction SilentlyContinue }catch{}
            try{remove-item "ScriptFolderLocation.txt" -Path "c:\yw-data\automate" -force -ErrorAction SilentlyContinue  }catch{}
            New-Item -Path c:\yw-data\automate -Name "PrinterName1.txt" -ItemType "file" -Value $printerName1 | out-null
            New-Item -Path c:\yw-data\automate -Name "ScriptFolderLocation.txt" -ItemType "file" -Value $ScriptFolderLocation | out-null
            (Get-Item c:\yw-data\automate\PrinterName1.txt).Attributes = "Hidden"
            (Get-Item c:\yw-data\automate\ScriptFolderLocation.txt).Attributes = "Hidden"
    
            Invoke-AsCurrentUser -ScriptBlock {
           
                $PrinterName1 = Get-Content "c:\yw-data\automate\PrinterName1.txt"
                $ScriptFolderLocation = Get-Content "c:\yw-data\automate\ScriptFolderLocation.txt"
                $config1 = Get-Content "$ScriptFolderLocation\PrinterConfiguration1.xml" | out-string
                Set-PrintConfiguration -PrinterName $printerName1 -PrintTicketXml $config1
                Remove-Item "c:\yw-data\automate\PrinterName1.txt" -force
                Remove-Item "c:\yw-data\automate\ScriptFolderLocation.txt" -force
    
             
            }
            send-Log -logText "Printer configuration 1 restored successfully" 
            send-CustomToastNofication -text "Configuration restored"
            send-CustomFinalToastNotification -SendToTeams
            
        }else{
            send-Log -logText "Printer $($printerName1) added successfully but failed to restore printer configuration" -type Warning -addTeamsMessage
            send-CustomToastNofication -text "Failed to restore configuration" -type Warning
            send-CustomFinalToastNotification -SendToTeams
            exit
    
        }


    }else{
        send-Log -logText "Either printer exists and action is ADD or printer does not exist and action is UPDATE. Skipping applying printer configuration"
    }

}



if ($env:PrinterName2 -and ($env:printingDefaults2 -notlike 'skip')){ #skip it action is skip or printer name is not set

    if ($env:action -eq 'add' -and $printer2Exists){ #skip if printer already exists with action ADD
        send-Log -logText "Skipping printer configuration for $($printerName2) as it already exists" -addDashes below
        send-CustomFinalToastNotification
    
    }elseif(($env:action -eq 'update' -and $printer2Exists) -or ($env:action -eq 'add' -and -not $printer2Exists)){ #proceed if action is UPDATE and printer exists or ADD and printer does not exist


        send-Log -logText "Restoring printer configuration for $($printerName2): $($env:printingDefaults2)" -addDashes Below
    
        $config = Get-Content "$ScriptFolderLocation\PrinterConfiguration2.xml" | out-string
    
        if ($config){
    
            #we can't read parent variables with invoke-ascurrentuser so we need to place them into file
            try{remove-item "PrinterName2.txt" -Path "c:\yw-data\automate" -force  -ErrorAction SilentlyContinue }catch{}
            try{remove-item "ScriptFolderLocation.txt" -Path "c:\yw-data\automate" -force -ErrorAction SilentlyContinue  }catch{}
            New-Item -Path c:\yw-data\automate -Name "PrinterName2.txt" -ItemType "file" -Value $printerName2 | out-null
            New-Item -Path c:\yw-data\automate -Name "ScriptFolderLocation.txt" -ItemType "file" -Value $ScriptFolderLocation | out-null
            (Get-Item c:\yw-data\automate\PrinterName2.txt).Attributes = "Hidden"
            (Get-Item c:\yw-data\automate\ScriptFolderLocation.txt).Attributes = "Hidden"
    
            Invoke-AsCurrentUser -ScriptBlock {
           
                $PrinterName2 = Get-Content "c:\yw-data\automate\PrinterName2.txt"
                $ScriptFolderLocation = Get-Content "c:\yw-data\automate\ScriptFolderLocation.txt"
                $config2 = Get-Content "$ScriptFolderLocation\PrinterConfiguration2.xml" | out-string
                Set-PrintConfiguration -PrinterName $printerName2 -PrintTicketXml $config2
                Remove-Item "c:\yw-data\automate\PrinterName2.txt" -force
                Remove-Item "c:\yw-data\automate\ScriptFolderLocation.txt" -force
    
             
            }
            send-Log -logText "Printer configuration 2 restored successfully" 
            send-CustomToastNofication -text "Configuration applied"
            send-CustomFinalToastNotification 
            
        }else{
            send-Log -logText "Printer $($printerName2) added successfully but failed to restore printer configuration" -type Warning -addTeamsMessage
            send-CustomToastNofication -text "Failed to restore configuration" -type Warning
            send-CustomFinalToastNotification -SendToTeams
            exit
    
        }


    }else{
        send-Log -logText "Either printer exists and action is ADD or printer does not exist and action is UPDATE. Skipping applying printer configuration"
    }

}

   
    
#endregion printer confiruation