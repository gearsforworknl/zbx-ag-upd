# Attempts to install the latest version of Zabbix

param (
  [string]$majorver = "5.4",
  [string]$minorver = "5.4.9",
  [string]$zbx_var = "2",
  [string]$zbx_path = "C:\Program Files\Zabbix Agent"
)
$base_url = "https://cdn.zabbix.com/zabbix/binaries/stable/"
$req_url_1 = $base_url + $majorver + "/" + $minorver +"/zabbix_agent-" + $minorver + "-windows-amd64-openssl.zip"
$req_url_2 = $base_url + $majorver + "/" + $minorver +"/zabbix_agent" + $zbx_var + "-" + $minorver + "-windows-amd64-openssl-static.zip"


### SERVICES
############################################################

$zbx_ag1 = $zbx_path + "\zabbix_agentd.exe"
$zbx_ag2 = $zbx_path + "\zabbix_agent2.exe"
if(Test-Path $zbx_ag1) {
    # Attempts to stop the Zabbix service on the Windows machine
    &$zbx_ag1 --stop
    # Attempts to uninstall the Zabbix agent on the Windows machine
    &$zbx_ag1 --uninstall
} else {
    # Attempts to stop the Zabbix service on the Windows machine
    &$zbx_ag2 --stop
    # Attempts to uninstall the Zabbix agent on the Windows machine
    &$zbx_ag2 --uninstall
}

############################################################


### FILE DIRECTORIES
############################################################

# Creates Zabbix folder again
mkdir "$zbx_path\downloads"

# Downloads from Zabbix.com to c:\zabbix
Invoke-WebRequest "$req_url_2" -outfile "$zbx_path\downloads\zabbix.zip"

# Imports ZIP
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)

    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

# Unzipping file to c:\zabbix
Unzip "$zbx_path\downloads\zabbix.zip" "$zbx_path\downloads\" 

# Sorts files in c:\zabbix
Move-Item "$zbx_path\downloads\bin\zabbix_agent2.exe" -Destination $zbx_path -Force

# Cleans up downloads folder
Remove-Item "$zbx_path\downloads" -Force -Recurse

# Get conf file
$conf = (Get-ChildItem $zbx_path -Filter *.conf).Name

# Attempts to install the agent with the config in c:\zabbix
&$zbx_path\zabbix_agent2.exe --config "$zbx_path\$conf" --install

# Attempts to start the agent
&$zbx_path\zabbix_agent2.exe --start

# Creates Zabbix folder again
mkdir "$zbx_path\downloads"

# Downloads from Zabbix.com to c:\zabbix
Invoke-WebRequest "$req_url_1" -outfile "$zbx_path\downloads\zabbix.zip"

# Unzipping file to c:\zabbix
Unzip "$zbx_path\downloads\zabbix.zip" "$zbx_path\downloads\"

# Sorts files in c:\zabbix
Move-Item "$zbx_path\downloads\bin\zabbix_sender.exe" -Destination $zbx_path -Force

# Cleans up downloads folder
Remove-Item "$zbx_path\downloads" -Force -Recurse
