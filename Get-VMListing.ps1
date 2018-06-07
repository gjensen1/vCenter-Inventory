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

    ForEach ($VM in $VMS) {
        $VMGuest = $VM | Get-VMGuest
        $NICs = $VM | Get-NetworkAdapter
        $VMDisks = Get-HardDisk -VM $VM
        ForEach ($NIC in $NICs) {
            $into = New-Object PSObject
            Add-Member -InputObject $into -MemberType NoteProperty -Name Host -Value $VM.vmHost.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name Folder -Value $VM.Folder.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name Cluster -Value $VM.Folder.Parent.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMName -Value $VM.Name
            Add-Member -InputObject $into -MemberType NoteProperty -Name GuestOS -Value $VM.guest.OSFullName
            Add-Member -InputObject $into -MemberType NoteProperty -Name vCPU -Value $VM.NumCPU
            Add-Member -InputObject $into -MemberType NoteProperty -Name MemoryGB -Value $VM.MemoryGB
            Add-Member -InputObject $into -MemberType NoteProperty -Name C_DriveSizeGB -Value $VMdisks.CapacityGB[0]
            Add-Member -InputObject $into -MemberType NoteProperty -Name PowerState -Value $VM.PowerState
            Add-Member -InputObject $into -MemberType NoteProperty -Name NICType -Value $Nic.Type
            Add-Member -InputObject $into -MemberType NoteProperty -Name MACAddress -Value $Nic.MacAddress
            Add-Member -InputObject $into -MemberType NoteProperty -Name Network -Value $Nic.NetworkName
            Add-Member -InputObject $into -MemberType NoteProperty -Name IP-Addresses -Value $VMGuest.IPAddress[0]
            Add-Member -InputObject $into -MemberType NoteProperty -Name Address-Type -Value $NIC.ExtensionData.AddressType
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsVersion -Value $VM.Guest.ExtensionData.ToolsVersion
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsVersionStatus -Value $VM.Guest.ExtensionData.ToolsVersionStatus
            Add-Member -InputObject $into -MemberType NoteProperty -Name VMToolsRunningStatus -Value $VM.Guest.ExtensionData.ToolsRunningStatus
            $Data += $into
        }
    }
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
