$packageName = 'rebis-grade-excel-importer'
$installerType = 'zip'
$url = 'https://github.com/ucoruh/rebis-grade-excel-importer/releases/download/v1.1/windows-binaries.tar.gz'
$url64 = $url
$checksum = 'FCAD092165D9F1FC085191DEE42BC7BD08C600CE1892B177327C6A868CEFA1E9'
$checksumType = 'sha256'
$toolsDir = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$installDir = 'C:\Program Files\Rebis\Grade Excel Importer'

Install-ChocolateyZipPackage $packageName $url $toolsDir $checksumType $checksum

# Create the installation directory
New-Item -ItemType Directory -Path $installDir | Out-Null

# Copy the application files to the installation directory
Copy-Item -Path "$toolsDir\lib\*" -Destination $installDir -Recurse

# Add the installation directory to the system PATH
$envPath = [Environment]::GetEnvironmentVariable('PATH', [EnvironmentVariableTarget]::Machine)
$envPath += ";$installDir"
[Environment]::SetEnvironmentVariable('PATH', $envPath, [EnvironmentVariableTarget]::Machine)
