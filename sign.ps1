param([string]$FilePath)
$ErrorActionPreference = 'Stop'

# Ensure local .NET SDK is on PATH (for dotnet tool install)
$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path $localDotnet) {
    $env:PATH = "$localDotnet;$env:PATH"
}

# Read signing parameters from metadata.json
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$meta = Get-Content (Join-Path $scriptDir 'metadata.json') -Raw | ConvertFrom-Json

$params = @{
    Endpoint                = $meta.Endpoint
    CodeSigningAccountName  = $meta.CodeSigningAccountName
    CertificateProfileName  = $meta.CertificateProfileName
    Files                   = $FilePath
    FileDigest              = 'SHA256'
    TimestampRfc3161        = 'http://timestamp.acs.microsoft.com'
    TimestampDigest         = 'SHA256'
}

# Apply credential exclusions from metadata
foreach ($cred in $meta.ExcludeCredentials) {
    $params["Exclude$cred"] = $true
}

Invoke-TrustedSigning @params
