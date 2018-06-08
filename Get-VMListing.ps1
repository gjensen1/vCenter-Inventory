<#
# +------------------------------------------------------+
# |        Load VMware modules if not loaded             |
# +------------------------------------------------------+
"Loading VMWare Modules"
$ErrorActionPreference="SilentlyContinue" 
if ( !(Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) ) {
    if (Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI' ) {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VMware, Inc.\VMware vSphere PowerCLI'
       
    } else {
        $Regkey = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\VMware, Inc.\VMware vSphere PowerCLI'
    }
    . (join-path -path (Get-ItemProperty  $Regkey).InstallPath -childpath 'Scripts\Initialize-PowerCLIEnvironment.ps1')
}
$ErrorActionPreference="Continue"
#>

# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\VMListing"

#*****************
# Get VC from User
#*****************
Function Get-VCenter {
    [CmdletBinding()]
    Param()
    #Prompt User for vCenter
    Write-Host "Enter the FQHN of the vCenter containing the target Hosts: " -ForegroundColor "Yellow" -NoNewline
    $Global:VCName = Read-Host 
}
#*******************
# EndFunction Get-VC
#*******************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    [CmdletBinding()]
    Param()
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds -WarningAction SilentlyContinue
}
#***********************
# EndFunction Connect-VC
#***********************

#*******************
# Disconnect vCenter
#*******************
Function Disconnect-VC {
    [CmdletBinding()]
    Param()
    "Disconnecting $Global:VCName"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false
}
#**************************
# EndFunction Disconnect-VC
#**************************

#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {
    [CmdletBinding()]
    Param()
    "Building Local folder structure" 
    If (!(Test-Path $Global:Folder)) {
        New-Item $Global:Folder -type Directory
        }
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#***********************
# Function Get-VMListing
#***********************
Function Get-VMListing {
    [CmdletBinding()]
    Param()
    "Getting Listing of Virtual Machines in $Global:VCName"
    $VMs= Get-VM *
    $Data = @()
    $Count = 1
    
    ForEach ($VM in $VMS) {
        Write-Progress -Id 0 -Activity 'Generating VM Details ' -Status "Processing $($count) of $($VMs.count)" -CurrentOperation $_.Name -PercentComplete (($count/$VMS.count) * 100)
        $VMGuest = $VM | Get-VMGuest
        $NICs = $VM | Get-NetworkAdapter
        $VMDisks = Get-HardDisk -VM $VM
        #ForEach ($NIC in $NICs) {
            $into = New-Object PSObject
            Add-Member -InputObject $into -MemberType NoteProperty -Name Host -Value $VM.vmHost.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name Folder -Value $VM.Folder.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name Cluster -Value $VM.Folder.Parent.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMName -Value $VM.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name GuestOS -Value $VM.guest.OSFullName
            Add-Member -InputObject $into -MemberType NoteProperty -Name vCPU -Value $VM.NumCPU
            Add-Member -InputObject $into -MemberType NoteProperty -Name MemoryGB -Value $VM.MemoryGB
            If ($VMDisks.count -eq 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD1-Name -Value $VMDisks.Name
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD1_DrvSizeGB -Value $VMdisks.CapacityGB
            }
            If ($VMDisks.count -gt 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD1-Name -Value $VMDisks.Name[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD1_DrvSizeGB -Value $VMdisks.CapacityGB[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD2-Name -Value $VMDisks.Name[1]
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD2_DrvSizeGB -Value $VMdisks.CapacityGB[1]
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD3-Name -Value $VMDisks.Name[2]
                Add-Member -InputObject $into -MemberType NoteProperty -Name HD3_DrvSizeGB -Value $VMdisks.CapacityGB[2]
            }
            Add-Member -InputObject $into -MemberType NoteProperty -Name PowerState -Value $VM.PowerState

            if ($nics.count -eq 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Type -Value $Nics.Type
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-MACAddress -Value $Nics.MacAddress
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Network -Value $Nics.NetworkName
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Adr-Type -Value $NICS.ExtensionData.AddressType
            }
            if ($nics.count -gt 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Type -Value $Nics.Type[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-MACAddress -Value $Nics.MacAddress[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Network -Value $Nics.NetworkName[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC1-Adr-Type -Value $NICS.ExtensionData.AddressType[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC2-Type -Value $Nics.Type[1]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC2-MACAddress -Value $Nics.MacAddress[1]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC2-Network -Value $Nics.NetworkName[1]
                Add-Member -InputObject $into -MemberType NoteProperty -Name NIC2-Adr-Type -Value $NICS.ExtensionData.AddressType[1]
            }
            if ($VMGuest.IPAddress.Count -eq 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name IP-Adr1 -Value $VMGuest.IPAddress    
            }
            if ($VMGuest.IPAddress.Count -gt 1){
                Add-Member -InputObject $into -MemberType NoteProperty -Name IP-Adr1 -Value $VMGuest.IPAddress[0]
                Add-Member -InputObject $into -MemberType NoteProperty -Name IP-Adr2 -Value $VMGuest.IPAddress[1]
            }
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsVersion -Value $VM.Guest.ExtensionData.ToolsVersion
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsVersionStatus -Value $VM.Guest.ExtensionData.ToolsVersionStatus
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsRunningStatus -Value $VM.Guest.ExtensionData.ToolsRunningStatus
            $Data += $into
            $into = $null
            $Count++
        #}
    }
Write-Progress -Id 0 -Activity 'Generating VM Details ' -Completed
$Data | Export-CSV -Path $Global:Folder\$Global:VCname-VMList.csv -NoTypeInformation
}

#**************************
# Function Convert-To-Excel
#**************************
Function Convert-To-Excel {
    [CmdletBinding()]
    Param()
   "Converting HostList from $Global:VCname to Excel"
    $workingdir = $Global:Folder+ "\*.csv"
    $csv = dir -path $workingdir

    foreach($inputCSV in $csv){
        $outputXLSX = $inputCSV.DirectoryName + "\" + $inputCSV.Basename + ".xlsx"
        ### Create a new Excel Workbook with one empty sheet
        $excel = New-Object -ComObject excel.application 
        $excel.DisplayAlerts = $False
        $workbook = $excel.Workbooks.Add(1)
        $worksheet = $workbook.worksheets.Item(1)

        ### Build the QueryTables.Add command
        ### QueryTables does the same as when clicking "Data » From Text" in Excel
        $TxtConnector = ("TEXT;" + $inputCSV)
        $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($Connector.name)


        ### Set the delimiter (, or ;) according to your regional settings
        ### $Excel.Application.International(3) = ,
        ### $Excel.Application.International(5) = ;
        $query.TextFileOtherDelimiter = $Excel.Application.International(5)

        ### Set the format to delimited and text for every column
        ### A trick to create an array of 2s is used with the preceding comma
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        ### Execute & delete the import query
        $query.Refresh()
        $query.Delete()

        ### Get Size of Worksheet
        $objRange = $worksheet.UsedRange.Cells 
        $xRow = $objRange.SpecialCells(11).ow
        $xCol = $objRange.SpecialCells(11).column

        ### Format First Row
        $RangeToFormat = $worksheet.Range("1:1")
        $RangeToFormat.Style = 'Accent1'

        ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
        $Workbook.SaveAs($outputXLSX,51)
        $excel.Quit()
    }
    ## To exclude an item, use the '-exclude' parameter (wildcards if needed)
    remove-item -path $workingdir 

}
#*****************************
# EndFunction Convert-To-Excel
#*****************************

#***************
# Execute Script
#***************

# Get Start Time
$startDTM = (Get-Date)

CLS
#$ErrorActionPreference="SilentlyContinue"

"=========================================================="
" "
Write-Host "Get CIHS credentials" -ForegroundColor Yellow
$Global:Creds = Get-Credential -Credential $null

Get-VCenter
Connect-VC
"----------------------------------------------------------"
Verify-Folders
"----------------------------------------------------------"
Get-VMListing
"----------------------------------------------------------"
# Write-to-CSV
"----------------------------------------------------------"
Convert-To-Excel
"----------------------------------------------------------"
Disconnect-VC
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
#Clean-Up

# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
