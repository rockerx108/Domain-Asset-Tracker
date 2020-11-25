Import-Module ActiveDirectory
#Computers OU
    $STRComp = Get-ADComputer -searchbase "<<OU PATH>>" -Filter *  | Select-Object name


foreach ($Comp in $STRComp)
{

    $DComp=$Comp.name

    if (Test-Connection -computername $DComp -count 2 -quiet)
    
    {$DComp=$Comp.name
    
        #HD Serial Number
            $HDSN = Get-WmiObject Win32_PhysicalMedia -Computername $DComp | Where-Object {$_.TAG -eq '\\.\PHYSICALDRIVE0'}

        #HD Model, Size
            $HDModel = Get-WmiObject -class "win32_diskdrive" -Computername $DComp 
            $HDSize = Get-WmiObject -class "win32_diskdrive" -Computername $DComp 
            
            #Find HD Manufacturer from Model Number
            Add more lines here once you find more Serial number to manufacturer Data
                   if ($HDModel.model -like 'SAM*') {$HDManufacturer = "Samsung"}
                       Elseif ($HDModel.model -like 'ST*') {$HDManufacturer = "Seagate"}
                       Elseif ($HDModel.model -like 'WD*') {$HDManufacturer = "Western Digital"}
                       Elseif ($HDModel.model -like 'HDD*') {$HDManufacturer = "Toshiba"}
                       Else {$HDManufacturer = $HDModel
                   }
                    
            #Convert HDSize to HDSizeGB                               
                Foreach ($Partition in $HDSize){  
                   $HDSizeGB = $Partition.Size/1GB  
                   $HDSizeGB = [math]::round($HDSizeGB, 2)  
                } 

#Windows Version Number
            $Winver = get-wmiobject -Class win32_operatingsystem -computername $Dcomp 

        #Computer SerialNumber
            $SN = get-wmiobject -class "win32_bios" -Computername $DComp 

        #NIC MACAddress
        #May not work if you have virtualization software on the computer
           $MAC = Get-WmiObject -class "win32_networkadapter" -ComputerName $DComp | where-object {($_.name -like '*Ethernet connection*') -or ($_.name -like '*gigabit*')} 

        #NIC IPAddress
        #May not work if you have virtualization software on the computer
            $IP = get-wmiobject -class "win32_networkadapterconfiguration" -computername $DComp | where {($_.Description -like "*Ethernet*") -or ($_.Description -like '*gigabit*') -and ($_.DHCPEnabled -eq "True")} | Select-Object DNSHostName, @{N="IPAddress"; E={$_.IpAddress[0]}}

        #Computer Manufacturer, Model, Name
            $COMPManufacturer = Get-WmiObject -class "win32_computersystem" -computerName $DComp
            $COMPModel = Get-WmiObject -class "win32_computersystem" -computerName $DComp
            $COMPHostname = Get-WmiObject -class "win32_computersystem" -ComputerName $DComp
                     
        # Create new PSobject for array values to properties
            $asset = New-Object PSobject -Property @{'Manufacturer' = $COMPManufacturer.Manufacturer
                'Model Number' = $COMPModel.Model
                'Machine Serial' = $SN.SerialNumber
                'HD Manufacturer' = $HDManufacturer
                'HD Model#' = $HDModel.model
                'HD Size' = $HDSizeGB
                'HD Serial' = $HDSN.SerialNumber.trim()
                'Hostname' = $COMPHostname.Name
                'MAC Address' = $MAC.MACAddress
                'Windows Version' = $winver.version
                'IP Address' = $IP.IPaddress
                } 

            #Add IPAddress to $asset Object
            $asset | Add-Member -MemberType NoteProperty -Name IP -Value $IP.IPAddress -force
                                                     
        #Create $path to out put .csv files into
            $Path =  '.\output.csv'

        #Ouput values of $asset into .csv
            $asset | Select-Object -property 'Hostname','Windows Version','IP Address','Model Number','Machine Serial','HD Manufacturer','MAC Address' | export-csv -path $Path -Append
    }


    else {$DComp | Out-File -filepath .\Failed.txt -append}
}
