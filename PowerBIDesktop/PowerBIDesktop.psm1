Function Get-PowerBIDesktopConnections {
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

    [CmdletBinding()]

    param(
        [string] 
        $Title
    )

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

    if ($title -ne $null) {
        $pbiProcesses | Where-Object Title -Like $Title
    }
    else {
        $pbiProcesses        
    }
    

}

Function Invoke-PowerBIDesktopDAXCommand {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true)]
        [string]
        $Command,
        [string]
        $Title = "*"
    )

    $sessions = Get-PowerBIDesktopConnections -Title $Title

    if ($sessions.length -eq 0) {
        if ($Title) {
            Write-Error -ErrorAction Stop -Message "Could not find any active Power BI connection with title <$title>" 
        }
        else {
            Write-Error -ErrorAction Stop -Message "Could not find any active Power BI connections. Is Power BI Running?"
        }
    }
    elseif ($sessions.length -gt 1) {
        if ($Title) {
            Write-Error -ErrorAction Stop -Message ("There is more than one active Power BI connection with title <$title>. Please be more specific in your filter",($sessions | Format-Table) -Join "`r`n") 
        } else {
            Write-Error -ErrorAction Stop -Message ("There is more than one active Power BI connection avaoilable. Please use -Title to specify which connection to use",($sessions | Format-Table) -Join "`r`n") 
        }
    }
    $session = $sessions[0]

    Write-Verbose "Connecting to Power BI connection $($session.Title) [$($session.DataSource)]"

    try {
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient")  
        $connection = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection
    } catch {
        Write-Error -ErrorAction Stop -Message "Could not create the connection. ADOMD.NET not installed. Please go to https://www.microsoft.com/en-us/download/details.aspx?id=52676 and download and install SQL_AS_ADOMD.msi"
    }

    try {
        $connection.ConnectionString = "Data Source=localhost:$($session.port)"
        $connection.open()
    } catch {
        Write-Error "Error connecting to $($session.Title) [$($connection.ConnectionString)]"
        throw $_
    }

    try {
        Write-Verbose "Executing the DAX command: $Command"
        $com = $connection.CreateCommand()
        $com.CommandText = "$Command"
        $adapter = New-Object -TypeName Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $com
        $dataset = New-Object -TypeName System.Data.DataSet

        $adapter.Fill($dataset)
        
    } catch {
        Write-Error "Error Executing DAX Statement: $Command"
        throw $_
    }
    
    $dataset
}
