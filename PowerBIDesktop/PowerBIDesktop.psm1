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
    # This used to be $_.Address:$_.Port but it appears that ADOMD.Net doesn't support addresses of type "::1:23456" for the data source parameter
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName DataSource -NotePropertyValue "localhost:$($_.Port)"}
    $pbiProcesses | ForEach-Object {Add-Member -InputObject $_ -NotePropertyName Title -NotePropertyValue $_.MainWindowTitle.Replace(" - Power BI Desktop", "")}
    
    # Remove unneeded information
    $pbiProcesses = $pbiProcesses | Select-Object Id, Title, DataSource, Address, Port

    if ($title) {
        $pbiProcesses | Where-Object Title -Like $Title
    }
    else {
        $pbiProcesses        
    }
    

}

Function Invoke-PowerBIDesktopCommand {
    <#
    .SYNOPSIS
    Run a command on a Power BI Desktop Analysis Server session.
    .Description
    Run a command on a Power BI Desktop Analysis Server session.
    This cmdlet allows you to pass a DAX, MDX or ASSL (XMLA) command on a running Power BI Desktop backend.
    Outputs a dataset containing the result of the command.
    Use -Title to specify the window title of the Power BI Desktop instance that you want to connect to, if there is more than one Power BI Desktop window open.

    Power BI Desktop internally runs queries against a modified version of SQL Server Analysis Services, and this cmdlet connects to it using ADOMD.Net.
    It creates a Dataset that connects to the query result, and then returns the first table from its Tables collection.

    For more information on DAX: https://msdn.microsoft.com/query-bi/dax/dax-queries
    For more information on MDX: https://docs.microsoft.com/en-us/sql/analysis-services/multidimensional-models/mdx/mdx-query-fundamentals-analysis-services
    For more information on ASSL (XMLA): https://docs.microsoft.com/en-us/sql/analysis-services/scripting/analysis-services-scripting-language-assl-for-xmla
    .PARAMETER Command
    The Command to run
    .PARAMETER Title
    The window title of the Power BI Desktop instance to run it against. 
    This can contain wild cards: "Fabrikam*" or "*eports*"

    This is not needed if there is only one Power BI Desktop instance running
    .EXAMPLE
    Retrieve the tables in the only running Power BI Dekstop instance


    #>

    [CmdletBinding()]

    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Command,
        [string]
        $Title
    )

    if ($title) {
        $sessions = Get-PowerBIDesktopConnections -Title $Title
    }
    else {
        $sessions = Get-PowerBIDesktopConnections
    }
    

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
            Write-Error -ErrorAction Stop -Message ("There is more than one active Power BI connection with title <$title>. Please be more specific in your filter", ($sessions | Format-Table) -Join "`r`n") 
        }
        else {
            Write-Error -ErrorAction Stop -Message ("There is more than one active Power BI connection avaoilable. Please use -Title to specify which connection to use", ($sessions | Format-Table) -Join "`r`n") 
        }
    }
    $session = $sessions[0]

    Write-Verbose "Connecting to Power BI connection $($session.Title) [$($session.DataSource)]"

    try {
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient")  
        $connection = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection
    }
    catch {
        Write-Error -ErrorAction Stop -Message "Could not create the connection. ADOMD.NET not installed. Please go to https://www.microsoft.com/en-us/download/details.aspx?id=52676 and download and install SQL_AS_ADOMD.msi or using nuget install Microsoft.AnalysisServices.AdomdClient.retail.amd64"
    }

    try {
        $connection.ConnectionString = "Data Source=$($session.DataSource)"
        $connection.open()
    }
    catch {
        Write-Error "Error connecting to $($session.Title) [$($connection.ConnectionString)]"
        throw $_
    }

    try {
        Write-Verbose "Executing the Power BI Desktop command: $Command"
        $com = $connection.CreateCommand()
        $com.CommandText = "$Command"
        $adapter = New-Object -TypeName Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $com
        $dataset = New-Object -TypeName System.Data.DataSet

        $adapter.Fill($dataset)
        
    }
    catch {
        Write-Error "Error Executing Power BI Desktop command: $Command"
        throw $_
    }
    
    $connection.close()

    $dataset
}

Function Get-PowerBIDesktopTables {
    <#
    .SYNOPSIS
    Gets the list of available tables in a Power BI Desktop Session and their descriptions
    .Description
    Gets the list of available tables in a Power BI Desktop Session and their descriptions
    If there is more than one Power BI Desktop Session active, you can specify the one you want to query with -Title
    There may be tables that are hidden from the normal Power BI interface. You can show them with -IncludeHidden
    .PARAMETER Title
    The window title of the Power BI Desktop instance to run it against. 
    This can contain wild cards: "Fabrikam*" or "*eports*"

    This is not needed if there is only one Power BI Desktop instance running
    .PARAMETER IncludeHidden
    This will include the tables that are hidden in the normal Power BI Desktop Interface
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]
        $Title = $null,
        # Parameter help description
        [Parameter(Mandatory = $false)]
        [Switch]
        $IncludeHidden
    )

    
    $dataset = Invoke-PowerBIDesktopCommand "SELECT * FROM `$SYSTEM.TMSCHEMA_TABLES" -Title $Title
    
    if ($IncludeHidden) {
        $dataset.Tables[0] | Select-Object Name, Description
    }
    else {
        $dataset.Tables[0] | where-object {$_.IsHidden -eq $false } | Select-Object Name, Description
    }
}

Function Read-PowerBIDesktopTable {
    <#
    .SYNOPSIS
    Reads the contents of the given PowerBI Desktop table and returns the result as a DataTable
    .Description
    Reads the contents of the given PowerBI Desktop table and returns the result as a DataTable
    .Parameter Table
    The table to read
    .PARAMETER Title
    The window title of the Power BI Desktop instance to run it against. 
    This can contain wild cards: "Fabrikam*" or "*eports*"

    This is not needed if there is only one Power BI Desktop instance running
    .Notes
    Currently does not correctly escape table names with a ' character. Please supply it as "''".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Table,
        [string]
        $Title = $null
    )

    $dataset = Invoke-PowerBIDesktopCommand -Command "EVALUATE ('$table')" -Title $Title
    $dataset.Tables[0]
}