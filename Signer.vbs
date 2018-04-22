'-----------------------------------------------------------------------------------------------------
'  This script will digitally sign any file dragged & dropped on to it with the private key
'  of the certificate called "Administrator".  The certificate must be signed by a trusted 
'  certificate authority.  In addition, the host running this script must be
'  using WSH 5.6 or later.  For example purposes, you can import the "CohoVineyard CA" 
'  CA certificate into your trusted root certification store, then import the 
'  "Administrator" Code Signing Certificate (signer.pfx) to use for digitally signing the scripts.  
'  Password to import the private key is "signer".
'
'
'  By default, clients will not check the digital signature of a script.  To enforce
'  script execution policies, clients will require the following registry modification:
'  Set HKLM\Software\Microsoft\Windows Script Host\Settings\TrustPolicy to a value of 2.
'  After this modification, clients will only be allowed to execute scripts that have
'  been digitally signed by a trusted code signing certificate issued from a trusted CA.  The easiest 
'  way to get clients to trust your code signing certificate is to import the certificate as a Trusted
'  Publisher on a client machine (this may require you to temporarily set the TrustPolicy registry
'  value on the client to 1), then use that client machine to create a Group Policy for Internet
'  Explorer Maintenance -> Security -> Authenticode Settings that mirrors the client's own policy.
'
'  Once you begin signing your scripts, if the signed script is modified in any way, the signature
'  will fail verification, and the script will not execute on clients configured to verify signatures.
'
'  Shawn Stugart
'====================================================================================================

Option Explicit
dim oFilesToSign, file, objSigner, i
set objSigner = WScript.CreateObject("Scripting.Signer")
set oFilesToSign = WScript.Arguments

For i = 0 to oFilesToSign.Count - 1
    file = oFilesToSign(i)
    objSigner.SignFile file, "Administrator"
Next
WScript.Echo "Signing Complete."
WScript.Quit