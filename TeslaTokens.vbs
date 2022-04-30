'Prerequesites:
'Windows
'Chilkat: http://www.chilkatsoft.com/downloads_ActiveX.asp

'Based on the work on those two websites
'- https://tesla-api.timdorr.com/api-basics/authentication
'- https://github.com/bntan/tesla-tokens

'Version 1.0 (2022-04-30)

Option Explicit
Dim fso, wsh, dict, fortuna, crypt, http, json, outFilePath, outFile, scriptdir
Dim code_verifier, code_challenge, state, lang, success
Dim code, data, response, URL, jsonRequestBody, resp, strValue
Dim LangFilePath, LangFile, English, Line, Pos, EndPos

Const Title = "TeslaTokens"
Set fso = CreateObject("Scripting.FileSystemObject")
Set wsh = CreateObject("WScript.Shell")
Set dict = CreateObject("Scripting.Dictionary")
Set fortuna = CreateObject("Chilkat_9_5_0.Prng")
Set crypt = CreateObject("Chilkat_9_5_0.Crypt2")
Set http = CreateObject("Chilkat_9_5_0.Http")
scriptdir = fso.GetParentFolderName(WScript.ScriptFullName)
outFilePath = Replace(WScript.ScriptFullName,".vbs",".txt")

'Read language from OS and read localisation file in dictionary
lang = ReadFromRegistry("HKEY_CURRENT_USER\Control Panel\International\LocaleName","en-US")
LangFilePath = scriptdir & "\TeslaTokens-" & lang & ".txt"
if fso.FileExists(LangFilePath) then
	Set LangFile = fso.OpenTextFile(LangFilePath,1)
	Do Until LangFile.AtEndOfStream
		Line = LangFile.ReadLine
		Pos = InStr(Line,"=")
		if Pos > 0 then
			dict.Add Left(Line,Pos-1),Mid(Line,Pos+1)
		end if
	Loop
else
	'Fallback to english as default
	English = True
end if

'Create random strings with length of 83 and 16 characters
code_verifier = fortuna.RandomString(83,1,1,1)
state = fortuna.RandomString(16,1,1,1)

'SHA256
crypt.HashAlgorithm = "sha256"
crypt.EncodingMode = "hex"
code_challenge = crypt.HashStringENC(code_verifier)

'URLSAFE64
crypt.EncodingMode = "base64url"
code_challenge = crypt.EncodeString(code_challenge,"ansi","base64")
code_challenge = replace(code_challenge,"=","")

MsgBox GetText("Welcome",""),,Title
MsgBox GetText("Introduction",""),,Title

'Make URL and call it in default browser. User has to log in his TESLA account
URL = "https://auth.tesla.com/oauth2/v3/authorize?audience=https%3A%2F%2Fownership.tesla.com%2F&client_id=ownerapi&code_challenge=" & code_challenge & "&code_challenge_method=S256&locale=en-US&prompt=login&redirect_uri=https%3A%2F%2Fauth.tesla.com%2Fvoid%2Fcallback&response_type=code&scope=openid+email+offline_access&state=" & state
wsh.Run URL

'After authentiation user must now enter the URL. Code will be extracted from the URL
code = InputBox(GetText("InputCode",""),Title)
if Len(Code) > 0 then
	Pos = InStr(code,"code=")
	if Pos > 0 then Pos = Pos + 5
	EndPos = InStr(Pos,code,"&")
	if EndPos > 0 then EndPos = EndPos
	code = mid(code,Pos,EndPos - Pos)
else
	WScript.Quit
end if

jsonRequestBody = "{ ""grant_type"": ""authorization_code"", ""client_id"": ""ownerapi"", ""code_verifier"": """ & code_verifier & """, ""code"": """ & code & """, ""redirect_uri"": ""https://auth.tesla.com/void/callback"" }"

http.SetRequestHeader "User-Agent",""
http.SetRequestHeader "x-tesla-user-agent",""
http.SetRequestHeader "X-Requested-With","com.teslamotors.tesla"
http.Accept = "application/json"
URL = "https://auth.tesla.com/oauth2/v3/token"

Set resp = http.PostJson2(url,"application/json",jsonRequestBody)
If (http.LastMethodSuccess = 0) Then
    MsgBox http.LastErrorText,,Title
    WScript.Quit
End If

if resp.StatusCode = 200 then
	Set outFile = fso.CreateTextFile(outFilePath, True)
	set json = CreateObject("Chilkat_9_5_0.JsonObject")
	success = json.Load(resp.BodyStr)
	json.EmitCompact = 0
	outFile.WriteLine(GetText("HeaderTokenFile",DateTime()))
	outFile.WriteLine(json.Emit())
	outFile.Close
	MsgBox GetText("Success",outFilePath),,Title
	wsh.Run outFilePath
else
	MsgBox GetText("Error",resp.StatusCode),,Title
end if

Function DateTime()
	DateTime = year(now()) & "-" & right("0" & month(now()),2) & "-" & right("0" & day(now()),2) & " " & right("0" & hour(now()),2) & ":" & right("0" & minute(now()),2) & ":" & right("0" & second(now()),2)
End Function

Function GetText(p, para)
	Dim NotFound
	if NOT English then
		if dict.Exists(p) then
			GetText = dict.Item(p)
			if len(para) > 0 then
				GetText = Replace(GetText,"###",para)
			end if
			NotFound = False
		else
			NotFound = True
		end if
	end if
	if English or NotFound then
		Select Case p
			Case "Welcome"
				GetText = "Welcome. This scripts gets the access tokens for your TESLA car. Since you can open this script with any editor you can be sure that none of your data will be grabbed."
			Case "Introduction"
				GetText = "After you press OK, the TESLA login window will open in your web browser. Please register there."
			Case "InputCode"
				GetText = "After you are logged in, 'Page Not Found' appears, but everything is fine. Please copy the URL line from your web browser and paste it here:"
			Case "Success"
				GetText = "Your tokens have been successfully read and are in file ###. This file will then be opened in your editor."
			Case "Error"
				GetText = "The TESLA website returned status code ###. Your tokens could not be read."
			Case "HeaderTokenFile"
				GetText = "Your access and refresh tokens (###):"
		End Select
	end if
End Function

Function ReadFromRegistry(strRegistryKey, strDefault)
    Dim value

    On Error Resume Next
    value = wsh.RegRead( strRegistryKey )
    if err.number <> 0 then
        readFromRegistry= strDefault
    else
        readFromRegistry=value
    end if
End Function
