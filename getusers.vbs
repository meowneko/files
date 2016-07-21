Const ADS_UF_ACCOUNTDISABLE = &H0003 
Const ADS_UF_LOCKOUT = &H0010 
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000 
Const ADS_UF_PASSWORD_EXPIRED = &H800000 
Const ADS_SECURE_AUTHENTICATION = &H1
Const ADS_USE_ENCRYPTION = &H2

computer = ""
user = ""
pass = ""

If Wscript.Arguments.Count = 1 Then
	computer =  WScript.Arguments.Item(0)
ElseIf Wscript.Arguments.Count = 3 Then
	computer = WScript.Arguments.Item(0)
	user = WScript.Arguments.Item(1)
	pass = WScript.Arguments.Item(2)
Else
	Wscript.Echo "usage: getusers.vbs <computer> (<user> <pass>)"
	Wscript.Quit 1
END If


Set objDSO = GetObject("WinNT:")
Set colAccounts = objDSO.OpenDSObject("WinNT://" & computer, user, pass, ADS_SECURE_AUTHENTICATION OR ADS_USE_ENCRYPTION)
colAccounts.Filter = Array("User")

Wscript.echo "HOST,USERNAME,ACCOUNT_DISABLED?,ACCOUNT_LOCKED?,PASS_SET_TO_EXPIRE?,PASS_EXPIRED?,PASS_LAST_CHANGED,BAD_PASS_ATTEMPTS,NAME,FULLNAME,DESC"
For Each objUser In colAccounts
	Dim info
	info = ",AccountDisabled: " & objUser.AccountDisabled
	info = info & ",IsAccountLocked: " & objUser.IsAccountLocked

	flag = objUser.Get("UserFlags")
	If flag AND ADS_UF_DONT_EXPIRE_PASSWD Then
		info = info & ",PassNoExpire"
	Else
		info = info & ",PassExpires"
	End If

	info = info & ",PasswordExpired:" & objUser.Get("PasswordExpired")

	intPasswordAge = objUser.Get("PasswordAge")
	intPasswordAge = intPasswordAge * -1 
	dtmChangeDate = DateAdd("s", intPasswordAge, Now)
	info = info & ",PasswordLastChanged: " & dtmChangeDate

	info = info & ",BadPasswordAttempts:" & objUser.Get("BadPasswordAttempts")
	info = info & ",Name:" & objUser.Get("Name")
	info = info & ",FullName:" & objUser.Get("FullName")
	info = info & ",Description:" & objUser.Get("Description")

	Wscript.Echo computer & "," & objUser.Name & info
Next
objDSO = nothing
colAccounts = nothing
