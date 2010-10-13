$varName="iResearch"
$varVal=Get-Location

[Environment]::SetEnvironmentVariable($varName, $varVal, "Machine")