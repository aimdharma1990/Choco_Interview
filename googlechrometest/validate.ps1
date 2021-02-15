$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$WebClient = New-Object System.Net.WebClient

#Getting the old Chrome verison from the nuspec package





$url = "https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise.msi"
$output = $scriptPath+'\chromeInstaller.msi' 

$WebClient.DownloadFile($url, $output)

[xml]$cn = Get-Content $scriptPath"\googlechrometest.nuspec"
$existingversion =  $cn.package.metadata.version
#Invoke-WebRequest -Uri $url -OutFile $output



#Getting the Msi version information
$windowsInstaller = New-Object -com WindowsInstaller.Installer
$pathToMSI = $scriptPath+"\chromeInstaller.msi"

$database = $windowsInstaller.GetType().InvokeMember(
	"OpenDatabase", "InvokeMethod", $Null,
	$windowsInstaller, @($pathToMSI, 0)
)





$q = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
$View = $database.GetType().InvokeMember(
	"OpenView", "InvokeMethod", $Null, $database, ($q)
)

$View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)
$record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null )
$newversion = $record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $record, 1 )

$view.GetType().InvokeMember("Close", "InvokeMethod", $null, $view, $null)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($view) | Out-Null
$database.GetType().InvokeMember("Commit", "InvokeMethod", $Null, $database, $Null)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($database) | Out-Null
$record=$null
$database=$null
if ( $newversion -ne $existingversion )
{

#delete the existing googlechrometes nupkg 


$FileName = $scriptPath+"\googlechrometest."+$existingversion+".nupkg"

if (Test-Path $FileName) {
  Remove-Item $FileName
}


#Then update nuspec file properties
$cn.package.metadata.version = $newversion
$cn.Save( $scriptPath+"\googlechrometest.nuspec")

 choco pack $scriptPath"\googlechrometest.nuspec"

$FileName = $scriptPath+"\googlechrometest."+$newversion+".nupkg"

if (Test-Path $FileName) {
  Write-host "Choco pack is success"
}

}

if (Test-Path $output) {
  Move-Item -Path $output -Destination $scriptPath\tools\googlechrome_x32.msi -Force
  $url = "https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi"
  $output = $scriptPath+'\tools\chromeInstaller_x64.msi' 
  $WebClient.DownloadFile($url, $output)  
}