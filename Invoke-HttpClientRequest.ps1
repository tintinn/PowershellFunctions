Add-Type -AssemblyName System.Net.Http
function Invoke-HTTPClientRequest {

    param(
    [Parameter (Mandatory = $true)] [String]$URI,
    [Parameter (Mandatory = $true)] [String]$Method,
    [Parameter (Mandatory = $false)] [Hashtable]$Headers,
    [Parameter (Mandatory = $false)] [HashTable]$Body,
    [Parameter (Mandatory = $false)] [int]$Timeout
    )

    $cookieContainer = [System.Net.CookieContainer]::new()

    $handler = [System.Net.Http.HttpClientHandler]::new();
    $handler.ClientCertificateOptions = [System.Net.Http.ClientCertificateOption]::Automatic;
    $handler.SslProtocols = [System.Security.Authentication.SslProtocols]::Tls12;

    $client = [System.Net.Http.HttpClient]::new($handler)

    if($Timeout -gt 0){
        $client.Timeout = New-TimeSpan -Seconds $Timeout
    }
    
    $request = [System.Net.Http.HttpRequestMessage]::new()
    $request.Method = $Method
    $request.RequestUri = $URI

    foreach($item in $Headers.GetEnumerator()){
        $request.Headers.Add($item.Name, $item.Value)
    }

    if($Body -ne $null){
        $encodedItems = @()
        foreach( $item in $Body.GetEnumerator()){
            $encodedItems += [System.Net.WebUtility]::UrlEncode($item.Name) + "=" + [System.Net.WebUtility]::UrlEncode($item.Value)
        }

        $content_string = ($encodedItems -join "&")
        Write-Host "Writing " $content_string.Length " bytes"
    
        $request.Content = [System.Net.Http.StringContent]::new( $content_string, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded");
        $request.Content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/x-www-form-urlencoded");
    }

    $HttpResponseMessageObject = $null
    $Response = $null
    try {
        $HttpResponseMessageObject = $client.SendAsync($request).GetAwaiter().GetResult()
        $HttpResponseMessageObject.EnsureSuccessStatusCode()
        $Response =  $HttpResponseMessageObject.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    } catch {
        $d = $(Get-Date)
        Write-Host "Failed at: $d"
        if($_.ErrorDetails.Message) {
            $Response = $_.ErrorDetails.Message
        } else {
            $Response = $_.Exception.ToString()
        }
        return $Response
    }

    return $Response
}
