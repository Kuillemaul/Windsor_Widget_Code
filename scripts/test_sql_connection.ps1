param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$false)][string]$Port = '',
    [Parameter(Mandatory=$true)][string]$Database,
    [Parameter(Mandatory=$true)][string]$Username,
    [Parameter(Mandatory=$true)][string]$Password,
    [Parameter(Mandatory=$true)][string]$OutputFile
)

$ErrorActionPreference = 'Stop'

function Write-Result([string]$Message, [int]$ExitCode) {
    $dir = Split-Path -Parent $OutputFile
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Set-Content -Path $OutputFile -Value $Message -Encoding UTF8
    exit $ExitCode
}

try {
    Add-Type -AssemblyName System.Data
    $builder = New-Object System.Data.Odbc.OdbcConnectionStringBuilder
    $builder["Driver"] = "ODBC Driver 18 for SQL Server"
    if ([string]::IsNullOrWhiteSpace($Port)) {
        $builder["Server"] = $Server.Trim()
    } else {
        $builder["Server"] = "{0},{1}" -f $Server.Trim(), $Port.Trim()
    }
    $builder["Database"] = $Database.Trim()
    $builder["Uid"] = $Username.Trim()
    $builder["Pwd"] = $Password
    $builder["Encrypt"] = "no"
    $builder["TrustServerCertificate"] = "yes"
    $builder["Connection Timeout"] = "5"

    $conn = New-Object System.Data.Odbc.OdbcConnection($builder.ConnectionString)
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT 1"
        [void]$cmd.ExecuteScalar()
    }
    finally {
        if ($conn.State -ne [System.Data.ConnectionState]::Closed) {
            $conn.Close()
        }
        $conn.Dispose()
    }

    Write-Result "Connection successful." 0
}
catch {
    $message = $_.Exception.Message
    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = $_ | Out-String
    }
    Write-Result ("Connection test failed.`r`n`r`n" + $message.Trim()) 1
}
