using namespace System.Security.Cryptography
using namespace Microsoft.Exchange.WebServices.Data
using namespace Microsoft.Identity.Client

Add-Type -AssemblyName  System.Security
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\Microsoft.Identity.Client\4.22.0\Microsoft.Identity.Client.dll"

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

####################
#Exchange Functions#
####################

function Send-EncryptedMailMessage {

    param(
    [Parameter (Mandatory = $true)] [ExchangeService]$ExchangeService,
    [Parameter (Mandatory = $true)] [String[]]$To,
    [Parameter (Mandatory = $true)] [String[]]$Certs,
    [Parameter (Mandatory = $false)] [String[]]$From,
    [Parameter (Mandatory = $false)] [String]$Subject,
    [Parameter (Mandatory = $false)] [String]$Body,
    [Parameter (Mandatory = $false)] [String[]]$Cc,
    [Parameter (Mandatory = $false)] [String]$AttachmentPath
    )

    $message = [EmailMessage]::new($ExchangeService)

    foreach($addy in $To){

            $message.ToRecipients.Add($addy)

        }

    $message.Subject = $Subject

    $message.ItemClass = "IPM.Note.SMIME";

    $randomString = -join ((48..57) + (97..122) | Get-Random -Count 32 | % {[char]$_})
    $boundary = "--T_$randomString"

    $MIMEMessage = New-Object system.Text.StringBuilder 

    $MIMEMessage.AppendLine("Content-type: multipart/mixed;") | Out-Null 
	$MIMEMessage.AppendLine("`tboundary=`"T_$randomString`"") | Out-Null 
    
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine("> This message is in MIME format. Since your mail reader does not understand") | Out-Null 
    $MIMEMessage.AppendLine("this format, some or all of this message may not be legible.") | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine($boundary) | Out-Null 
    
    $randomString = -join ((48..57) + (97..122) | Get-Random -Count 32 | % {[char]$_})
    $bBoundary = "--T_$randomString"

    $MIMEMessage.AppendLine("Content-type: multipart/alternative;") | Out-Null 
	$MIMEMessage.AppendLine("`tboundary=`"T_$randomString`"") | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine($bBoundary) | Out-Null 
    
    $MIMEMessage.AppendLine("Content-type: text/plain;") | Out-Null 
	$MIMEMessage.AppendLine("`tcharset=`"UTF-8`"") | Out-Null 
    $MIMEMessage.AppendLine("Content-transfer-encoding: 7bit") | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine($Body) | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.Append($bBoundary) | Out-Null 
    $MIMEMessage.AppendLine("--") | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
    
    if( $AttachmentPath -ne "") {
        $MIMEMessage.AppendLine($boundary) | Out-Null 

        $FileAttachment = Get-Item -Path $AttachmentPath

        $MIMEMessage.AppendLine("Content-type: text/plain; name=`"$($FileAttachment.Name)`";") | Out-Null 
        $MIMEMessage.AppendLine("Content-disposition: attachment;") | Out-Null 
	    $MIMEMessage.AppendLine("`tfilename=`"$($FileAttachment.Name)`"") | Out-Null 
        $MIMEMessage.AppendLine("Content-transfer-encoding: base64") | Out-Null 
        $MIMEMessage.AppendLine() | Out-Null 
        $MIMEMessage.AppendLine() | Out-Null 

        [Byte[]] $file_bytes = [System.IO.File]::ReadAllBytes((Convert-Path $FileAttachment.PSPath))
        $file_base64 = [Convert]::ToBase64String($file_bytes);
        $MIMEMessage.AppendLine($file_base64) | Out-Null 
        $MIMEMessage.Append($boundary) | Out-Null 
        $MIMEMessage.AppendLine("--") | Out-Null 
        $MIMEMessage.AppendLine() | Out-Null 
    }

    [Byte[]] $mm_bytes = [System.Text.Encoding]::ASCII.GetBytes($MIMEMessage.ToString())

    $content =  New-Object -TypeName Pkcs.ContentInfo -ArgumentList(,$mm_bytes) 

    #encrypt message
    $cmsrecipient = New-Object Pkcs.CmsRecipientCollection

    foreach($key in $cert){
        $cmsrecipient.Add(
            (New-Object Pkcs.CmsRecipient $key)
        )
    }

    $enveloped = New-Object Pkcs.EnvelopedCms $content 
    $enveloped.Encrypt($cmsrecipient)
    [Byte[]] $encrypted = $enveloped.Encode()

    $encrypted_base64 = [Convert]::ToBase64String($encrypted);
    
    $MIMEMessage = New-Object System.Text.StringBuilder 

    $MIMEMessage.AppendLine("MIME-Version: 1.0") | Out-Null 
    $MIMEMessage.AppendLine("Content-Type: application/pkcs7-mime; smime-type=enveloped-data;") | Out-Null 
    $MIMEMessage.AppendLine("`tname=`"smime.p7m`"") | Out-Null
    $MIMEMessage.AppendLine("Content-Disposition: attachment; filename=`"smime.p7m`"") | Out-Null
    $MIMEMessage.AppendLine("Content-Transfer-Encoding: base64") | Out-Null
    $MIMEMessage.AppendLine() | Out-Null 
    $MIMEMessage.AppendLine($encrypted_base64) | Out-Null 
    $MIMEMessage.AppendLine() | Out-Null 
     
    [byte[]] $byteContent =  [System.Text.Encoding]::ASCII.GetBytes($MIMEMessage.ToString())
    
    $message.MimeContent = New-Object -TypeName MimeContent -ArgumentList [System.Text.Encoding]::ASCII.HeaderName, $byteContent
    
    $mailbox = [Mailbox]::new($From)
    $folderId = [FolderId]::new([WellKnownFolderName]::Drafts, $mailbox)
    $message.Save($folderId)
    $message = [EmailMessage]::Bind($ExchangeService, $message.Id)
    $message.Send()

    #$message.SendAndSaveCopy()

    
}

function New-ExchangeObject {

    [OutputType([ExchangeService])]
    param(
    [Parameter (Mandatory = $false)] [String]$URI = "https://outlook.office365.com/EWS/Exchange.asmx",
    [Parameter (Mandatory = $true)] [String]$ClientId,
    [Parameter (Mandatory = $true)] [String]$ClientSecret,
    [Parameter (Mandatory = $true)] [String]$TenantId,
    [Parameter (Mandatory = $true)] [String]$UserId,
    [Parameter (Mandatory = $false)] [String[]]$ClientScopes = [String[]] ( "https://outlook.office365.com/.default" )

    )

    #Microsoft Authentication Library
    $CCA =  [ConfidentialClientApplicationBuilder]::Create($ClientID).WithClientSecret($ClientSecret).WithTenantId($TenantId).Build()

    #Request a token
    $Result = $CCA.AcquireTokenForClient($ClientScopes).ExecuteAsync();

    $ExchangeService = New-Object ExchangeService
    $ExchangeService.Url = New-Object Uri($URI)
    #$ExchangeService.Credentials = New-Object -TypeName ExchangeCredentials.OAuthCredentials -ArgumentList $Result.Result.AccessToken
    $ExchangeService.Credentials = [OAuthCredentials]::new($Result.Result.AccessToken)
    $ExchangeService.ImpersonatedUserId = New-Object -TypeName ImpersonatedUserId([ConnectingIdType]::SmtpAddress, $UserId)

    return $ExchangeService

}

##############
#AD Functions#
##############

function Get-UserADPublicKey {

    param(
    [Parameter (Mandatory = $true)] [String]$Filter,
    [Parameter (Mandatory = $false)] [String]$Attribute = "cn",
    [Parameter (Mandatory = $false)] [String]$CertAttribute = "usercertificate"
    )

    $search = [adsisearcher]"(&(ObjectCategory=Person)(ObjectClass=User)($Attribute=$Filter*))"

    $users = $search.FindAll()

    $certBytes = $users[0].Properties[$CertAttribute][0]

    $cert = new-object X509Certificates.X509Certificate2 -ArgumentList (,$certBytes)

    return $cert

}

