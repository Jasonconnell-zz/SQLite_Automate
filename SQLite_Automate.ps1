<#
.SYNOPSIS Get local database information
.DESCRIPTION Used to pull system information from the Automate SQLite databases.
.Example Get-AutomateReboot
.Example Get-AutomatePatching
.AUTHOR Jason Connell
#>


Function Get-AutomatePatching {

    try{
        Add-Type -Path "C:\Windows\ltsvc\System.Data.SQLite.dll"
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=C:\Windows\LTSvc\Databases\Patching.db"
        $con.Open()

        $sql = $con.CreateCommand()
        $sql.CommandText = "SELECT * FROM ComputerPatchingStats"


        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
        $data = New-Object System.Data.DataSet
        [void]$adapter.Fill($data)
        $sql.Dispose()
        $con.Close()
            return $data.Tables.rows}
    Catch{
            return 'Failed to connect to local database'
          }
}


Function Get-AutomateReboot {

    try{
        Add-Type -Path "C:\Windows\ltsvc\System.Data.SQLite.dll"
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=C:\Windows\LTSvc\Databases\Patching.db"
        $con.Open()

        $sql = $con.CreateCommand()
        $sql.CommandText = "SELECT * FROM RebootPolicies"


        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
        $data = New-Object System.Data.DataSet
        [void]$adapter.Fill($data)
        $sql.Dispose()
        $con.Close()
            return $data.Tables.rows}
    Catch{
            return 'Failed to connect to local database'
          }
}


Function Get-AutomateSoftwarePolicies {

    try{
        Add-Type -Path "C:\Windows\ltsvc\System.Data.SQLite.dll"
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
        $con.ConnectionString = "Data Source=C:\Windows\LTSvc\Databases\Patching.db"
        $con.Open()

        $sql = $con.CreateCommand()
        $sql.CommandText = "SELECT * FROM InstallSoftwarePolicies"


        $adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
        $data = New-Object System.Data.DataSet
        [void]$adapter.Fill($data)
        $sql.Dispose()
        $con.Close()
            return $data.Tables.rows}
    Catch{
            return 'Failed to connect to local database'
          }
}
