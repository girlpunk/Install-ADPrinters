#requires -version 2.0

<#
    .SYNOPSIS
        Install printers from active directory.
    .DESCRIPTION
        The PowerShell script which can be used to set the default printer.
    .EXAMPLE
        C:\PS> C:\Script\Install-ADPrinters.ps1

        This command shows how to set install printers from Active Directory.

#>

# Get rid of existing printers
Get-WMIObject Win32_Printer | foreach{$_.delete()}

# Get the OU for this computer
$ComputerOU = Get-ADOrganizationalUnit -Identity $(($adComputer = Get-ADComputer -Identity $env:COMPUTERNAME).DistinguishedName.SubString($adComputer.DistinguishedName.IndexOf("OU=")))

# Connect to active directory
$root = New-Object system.DirectoryServices.DirectoryEntry
# Search active directory for printers
$search = New-Object system.DirectoryServices.DirectorySearcher($root)
$search.PageSize = 1000
$search.Filter = "(objectCategory=printQueue)"
$search.SearchScope = "Subtree"
$search.SearchRoot = $root
# Get the details of each printer
$colProplist = "Name","UNCName","PortName","PrintShareName","ShortServerName"
foreach ($i in $colProplist){$search.PropertiesToLoad.Add($i) | out-null }
$RawADPrinters = $search.FindAll()

# List of unique print servers
$PrintServers = @()

foreach ($RawADPrinter in $RawADPrinters) {
    if ($PrintServers -notcontains $RawADPrinter.Properties.shortservername) {
        $PrintServers +=           $RawADPrinter.Properties.shortservername
    }
}

# List of unique printers
$ADPrinters = @()

foreach ($PrintServer in $PrintServers) {
    foreach ($Printer in Get-Printer -Computer $PrintServer | ? Shared -eq $True) {
        if ($ADPrinters -notcontains $Printer) {
            $ADPrinters +=           $Printer
        }
    }
}

# For each printer
foreach ($Printer in $ADPrinters) {
    # Search the printer's comment for the OU of this printer, with an asterisk on the end
    if ($Printer.Comment.Split("|") -Contains $ComputerOU+"*") {
        # Found it, installing printer on this PC
        Add-Printer -ConnectionName ("\\"+$Printer.ComputerName+"\"+$Printer.Name)
        # The asterisk is used to signify that this should be the default printer for computers in this OU
        # There's no direct PowerShell method for setting the default printer, so we have to do this through WMI
        # Get a WMI Instance of the printer we just installed
        $WMIPrinter = Get-WmiObject -Class Win32_Printer | Where{$_.Name -eq ("\\"+$Printer.ComputerName+"\"+$Printer.Name)}
        # Set it as default
        $WMIPrinter.SetDefaultPrinter()
    # Search the printer's comment for the OU of this printer, without an asterisk at the end
    } else if ($Printer.Comment.Split("|") -Contains $ComputerOU) {
        # Found it, installing printer on this PC
        Add-Printer -ConnectionName ("\\"+$Printer.ComputerName+"\"+$Printer.Name)
    }
}


