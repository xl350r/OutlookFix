strComputer = "."
strRegPathSuffix = "\Software\Microsoft\Office\15.0\Common\Identity"
strRegValueName = "Version"
intRegValueDec = "1"
Const HKEY_USERS = &H80000003
 
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = ""
oReg.EnumKey HKEY_USERS, strKeyPath, arrSubKeys
 
For Each subkey In arrSubKeys
    'wscript.echo subkey
    'We only want to do something if the subkey does not contain any of the following: .DEFAULT or S-1-5-18 or S-1-5-19 or S-1-5-20 or _Classes
    If NOT ((InStr(1,subkey,".DEFAULT",1) > 0) OR (InStr(1,subkey,"S-1-5-18",1) > 0) OR (InStr(1,subkey,"S-1-5-19",1) > 0) OR (InStr(1,subkey,"S-1-5-20",1) > 0) OR (InStr(1,subkey,"_Classes",1) > 0)) Then
	'Create desired registry key/value
	strKeyPath = subkey & strRegPathSuffix
	'wscript.echo strKeyPath
	oReg.CreateKey HKEY_USERS, strKeyPath
	oReg.SetDWORDValue HKEY_USERS, strKeyPath, strRegValueName, intRegValueDec
    End If
Next