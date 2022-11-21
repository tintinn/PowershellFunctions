using namespace System.Net
Add-Type -AssemblyName System.Net.Http
function Invoke-HTTPClientRequest {

    param(
    [Parameter (Mandatory = $true)] [String]$URI,
    [Parameter (Mandatory = $true)] [String]$Method,
    [Parameter (Mandatory = $false)] [Hashtable]$Headers,
    [Parameter (Mandatory = $false)] [HashTable]$Body,
    [Parameter (Mandatory = $false)] [int]$Timeout
    )

    $cookieContainer = [CookieContainer]::new()

    $handler = [Http.HttpClientHandler]::new();
    $handler.ClientCertificateOptions = [Http.ClientCertificateOption]::Automatic;
    $handler.SslProtocols = [System.Security.Authentication.SslProtocols]::Tls12;

    $client = [Http.HttpClient]::new($handler)

    if($Timeout -gt 0){
        $client.Timeout = New-TimeSpan -Seconds $Timeout
    }
    
    $request = [Http.HttpRequestMessage]::new()
    $request.Method = $Method
    $request.RequestUri = $URI

    if($Headers -ne $null){
        foreach($item in $Headers.GetEnumerator()){
            $request.Headers.Add($item.Name, $item.Value)
        }
    }

    if($Body -ne $null){
        $encodedItems = @()
        foreach( $item in $Body.GetEnumerator()){
            $encodedItems += [WebUtility]::UrlEncode($item.Name) + "=" + [WebUtility]::UrlEncode($item.Value)
        }

        $content_string = ($encodedItems -join "&")
        Write-Host "Writing " $content_string.Length " bytes"
    
        $request.Content = [Http.StringContent]::new( $content_string, [System.Text.Encoding]::UTF8, "application/x-www-form-urlencoded");
        $request.Content.Headers.ContentType = [Http.Headers.MediaTypeHeaderValue]::Parse("application/x-www-form-urlencoded");
    }

    #$HttpResponseMessageObject = $null
    $Response = $null
    try {
        $HttpResponseMessageObject = $client.SendAsync($request).GetAwaiter().GetResult()
        #$HttpResponseMessageObject.EnsureSuccessStatusCode()|  Out-Null
        $Response =  $HttpResponseMessageObject.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    } catch {
        $d = $(Get-Date)
        Write-Host "Failed at: $d"
        $Response = $_.Exception
    }

    return $Response
}
