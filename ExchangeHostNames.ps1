#Getting Exchange FQDNs from configured URLs

#Local Server Name
try {
 $ExchangeServer = (Get-ExchangeServer $env:computername).Name
}
catch {
}

#Autodiscover
try {
 $AutodiscoverFQDN = ((Get-ClientAccessService -Identity $ExchangeServer).AutoDiscoverServiceInternalUri.Host).ToLower()
 [array]$CertNames += $AutodiscoverFQDN 
}
catch {
}

#Outlook Anywhere
try {
 $OAExtFQDN = ((Get-OutlookAnywhere -Server $ExchangeServer).ExternalHostname.Hostnamestring).ToLower()
 [array]$CertNames += $OAExtFQDN

 $OAIntFQDN = ((Get-OutlookAnywhere -Server $ExchangeServer).Internalhostname.Hostnamestring).ToLower()
 [array]$CertNames += $OAIntFQDN
}
catch {
}

#OAB
try {
 $OABExtFQDN = ((Get-OabVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $OABExtFQDN

 $OABIntFQDN = ((Get-OabVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $OABIntFQDN 
}
catch {
}

#ActiveSync
try {
 $EASIntFQDN = ((Get-ActiveSyncVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $EASIntFQDN

 $EASExtFQDN = ((Get-ActiveSyncVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $EASExtFQDN
}
catch {
}

#EWS
try {
 $EWSIntFQDN = ((Get-WebServicesVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $EWSIntFQDN

 $EWSExtFQDN = ((Get-WebServicesVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $EWSExtFQDN
}
catch {
}

#ECP
try {
 $ECPIntFQDN = ((Get-EcpVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $ECPIntFQDN

 $ECPExtFQDN = ((Get-EcpVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $ECPExtFQDN
}
catch {
}

#OWA
try {
 $OWAIntFQDN =  ((Get-OwaVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $OWAIntFQDN

 $OWAExtFQDN = ((Get-OwaVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $OWAExtFQDN
}
catch {
}

#MAPI
try {
 $MAPIIntFQDN = ((Get-MapiVirtualDirectory -Server $ExchangeServer).Internalurl.Host).ToLower()
 [array]$CertNames += $MAPIIntFQDN

 $MAPIExtFQDN = ((Get-MapiVirtualDirectory -Server $ExchangeServer).ExternalUrl.Host).ToLower()
 [array]$CertNames += $MAPIExtFQDN 
}
catch {
}

#Make FQDNs unique
try {
 $CertNames = $CertNames | select –Unique
}
catch {
}

write-host "
Autodiscover Hostname: $AutodiscoverFQDN 
Outlook Anywhere Hostname (Internal): $OAIntFQDN
Outlook Anywhere Hostname (External): $OAExtFQDN
ActiveSync Hostname (Internal): $EASIntFQDN
ActiveSync Hostname (External): $EASExtFQDN
OAB Hostname (Internal): $OABIntFQDN 
OAB Hostname (External): $OABExtFQDN
EWS Hostname (Internal): $EWSIntFQDN
EWS Hostname (External): $EWSExtFQDN
ECP Hostname (Internal): $ECPIntFQDN
ECP Hostname (External): $ECPExtFQDN
OWA Hostname (Internal): $OWAIntFQDN
OWA Hostname (External): $OWAExtFQDN
MAPI Hostname (Internal): $MAPIIntFQDN
MAPI Hostname (External): $MAPIExtFQDN
"
write-host "
SANs needed for Certificate:
"
$CertNames

write-host "
Use this Hostname as Common Name (CN): $OWAExtFQDN
"