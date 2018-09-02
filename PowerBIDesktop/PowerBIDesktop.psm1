Function Get-PowerBIDesktopConnections 
{
<#
.SYNOPSIS
Get the connection data for currently running Power BI Desktop sessions
.DESCRIPTION
Polls the currently running Power BI Desktop Sessions to find out what their internal
connection ports are, and returns that information
.PARAMETER Title
The Window Title of the connection you're searching for. Accepts Wildcards such as "*data*"
.EXAMPLE
C:\PS> Get-PowerBIDesktopConnections | ft

   Id Title                      DataSource Address  Port
   -- -----                      ---------- -------  ----
84096 Fabrikam Processes         ::1:51125  ::1     51125
84664 Northwind Sales Monitoring ::1:61248  ::1     61248

.EXAMPLE
C:\PS> Get-PowerBIDesktopConnections -Title "Fabrikam*"


Id         : 84096
Title      : Fabrikam Processes
DataSource : ::1:51125
Address    : ::1
Port       : 51125

#>
    # Get the power BI Desktop Processes
    $pbiProcesses = Get-Process | 
            Where-Object ProcessName -eq "PBIDesktop" |  
            Where-Object MainWindowTitle -ne ""
    
    # Each process performs a local connection to MS Analysis Services in the background.
    # We get the ports on loopback connections. 
    $pbiPorts = Get-NetTCPConnection -OwningProcess @($pbiProcesses | Select-Object -ExpandProperty Id) |
            Where-Object State -eq "Established" | 
            Where-Object {$_.LocalAddress -eq $_.RemoteAddress}  | 
            Select-Object OwningProcess, RemoteAddress, RemotePort | 
            Sort-Object {$_.OwningProcess} | 
            Get-Unique -AsString

    # Now we decorate the process objects:
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName Port -NotePropertyValue ($pbiports | Where-Object OwningProcess -eq $_.Id).RemotePort}
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName Address -NotePropertyValue ($pbiports | Where-Object OwningProcess -eq $_.Id).RemoteAddress}
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName DataSource -NotePropertyValue ($_.Address, ":", $_.Port -Join "")}
    #$pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName ConnectionString -NotePropertyValue ($_.Address, ":", $_.Port -Join "")}
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName Title -NotePropertyValue $_.MainWindowTitle.Replace(" - Power BI Desktop", "")}
    
    # Remove unneeded information
    $pbiProcesses = $pbiProcesses | Select-Object Id, Title, DataSource, Address, Port

    $pbiProcesses

}

