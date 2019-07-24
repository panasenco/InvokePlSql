<#
.Synopsis
    PowerShell wrapper around sql.exe (SqlCl) command-line tool.
.Description
    Uses sql.exe that ships as part of SQL Developer to execute a PL/SQL command or script, return the output
    as objects, and display helpful error messages.
.Parameter ConnectString
    The connection string to use. Set the environment variable DEFAULTCONNECTION to not have to provide this every time.
.Parameter Query
    Query to execute - can be passed through the pipeline.
.Parameter Scalar
    Set this switch to return a single value from each query.
.Notes
    sql.exe from SQL Developer (in sqldeveloper/bin) must be in PATH.
    TODO: Implement -Update switch to return number of rows affected by each query.
#>
function Invoke-PlSql {
[CmdletBinding()]
    param (
        [Parameter(Position=0)]
        [string] $ConnectString = $env:DEFAULTCONNECTION,
        [Parameter(ValueFromPipeline=$true)]
        [string] $Query,
        [switch] $Scalar,
        [switch] $WhatIf
    )
    begin {
        $QueryString = ''
    }
    process {
        # Accumulate query in the query string
        $QueryString += ($Query -replace "`r`n","`n" -replace "$","`n")
    }
    end {
        # Create temporary login.sql file
        "SET FEEDBACK 1",
        "SET SQLFORMAT CSV" | 
        Set-Content "$env:TEMP\login.sql"
        # Add temp directory to SQLPATH
        $OldSqlPath = $env:SQLPATH
        $env:SQLPATH = $env:TEMP
        # Strip blank lines from query
        $QueryString = $QueryString  -replace '(?<=\n|^)\s*\n'
        if ($WhatIf) { return $QueryString }
        # Invoke the query
        $Output = $QueryString | sql -S $ConnectString
        # Trim leading and trailing blank strings in output
        $Output = $Output -join "`n" -replace '^\n*' -replace '\n*$' -split "`n"
        # Reset SQLPATH
        $env:SQLPATH=$OldSqlPath
        # Check for blank output
        if ([string]::IsNullOrWhiteSpace($Output)) {
            return
        }
        # Check for errors
        if ($Output[0] -match '^Error starting at .*') {
            throw ($Output -join "`n")
        }
        # Check for no rows
        if ($Output[-1] -eq 'no rows selected') {
            return
        }
        # Drop the last two rows: A blank row and an "N rows selected." row
        $Output = $Output | Select-Object -SkipLast 2
        # Check for rows with odd number of double quotes (split by double quotes results in even # of pieces)
        $UnterminatedRows = $Output | where {($_ -split '"').Count % 2 -eq 0}
        if ($UnterminatedRows.Count -gt 0) {
            Write-Warning -Message ("Below output contains unterminated strings - sanitize query with " +
                "REGEXP_REPLACE(value, '[[:cntrl:]]'):`n$($UnterminatedRows -join ""`n"")")
        }
        # Handle duplicate column headers
        $HeaderCounts = [ordered]@{}
        $Header = $Output[0] -split '","' -replace '"','' | foreach {
            $HeaderCounts[$_]++
            "$_$(if($HeaderCounts[$_] -gt 1){$HeaderCounts[$_]})"
        }
        # Get the resulting object array
        $Result = $Output | Select -Skip 1 | ConvertFrom-Csv -Header $Header
        # Handle null output
        if ($Result -eq $null) {
            $Result = '' | select -Property $Header
        }
        # Process switches
        if ($Scalar) {
            ($Result[0].psobject.properties | select -First 1).Value
        } else {
            $Result
        }
    }
}
