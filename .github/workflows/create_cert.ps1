$cert = New-SelfSignedCertificate -CertStoreLocation "cert:\CurrentUser\My" -Subject "CN=SlideSCITemp" -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider"
$password = ConvertTo-SecureString -String "password" -Force -AsPlainText
# Move PFX to the project directory specifically
$pfxPath = Join-Path $env:GITHUB_WORKSPACE "SlideSCI\SlideSCI_Temp.pfx"
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password
