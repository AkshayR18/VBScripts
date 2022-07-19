' GetADGroupMembers.vbs
' Sample VBScript to Get List of AD Group Members.
' CMD Usage: 
'     CScript <vbscript file path> <groupName>
' Ex: CScript C:\Scripts\GetADGroupMembers.vbs "Domain Admins"
' Author: http://www.morgantechspace.com/
' ------------------------------------------------------' 

Dim groupName,strMember
Dim objGroup,objMember

if Wscript.arguments.count = 0 then
    Wscript.echo "Invalid input parameters"
    Wscript.echo "   "
    Wscript.echo "Script Usage:"
    Wscript.echo "----------------------------------------"
    Wscript.echo "CScript <vbscript file path> <groupName>"
    Wscript.echo "---------------------------------------"
    Wscript.echo "Ex: CScript C:\Scripts\GetADGroupMembers.vbs ""Domain Admins"" "
    Wscript.echo "---------------------------------------"
    WScript.quit
else
	
  ' Get the group name from command line parameter
    groupName = WScript.Arguments(0)

end if

' Get the distinguished name of the group
Set objGroup = GetObject("LDAP://" & GetDN(groupName))

' List the member’s full name in the group
For Each strMember in objGroup.Member
    Set objMember =  GetObject("LDAP://" & strMember)
    Wscript.Echo objMember.CN
Next

WScript.quit
' Active Directory Group Members listed successfully using VBScript


'****************Function to Get DN of group****************
' 
Function GetDN(groupName)

Dim objRootDSE, adoCommand, adoConnection
Dim varBaseDN, varFilter, varAttributes
Dim adoRecordset

Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Search entire Active Directory domain.
Set objRootDSE = GetObject("LDAP://RootDSE")

varDNSDomain = objRootDSE.Get("defaultNamingContext")
varBaseDN = "<LDAP://" & varDNSDomain & ">"


' Filter on group objects.
varFilter = "(&(objectClass=group)(|(cn="& groupName &")(name="& groupName &")))"

' Comma delimited list of attribute values to retrieve.
varAttributes = "distinguishedname"

' Construct the LDAP syntax query.
strQuery = varBaseDN & ";" & varFilter & ";" & varAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 1000
adoCommand.Properties("Timeout") = 20
adoCommand.Properties("Cache Results") = False

' Run the query.
Set adoRecordset = adoCommand.Execute


IF(adoRecordset.EOF<>True) Then
   GetDN=adoRecordset.Fields("distinguishedname").value
Else 
   'No group found 
End if

' close ado connections.
adoRecordset.Close
adoConnection.Close

End Function

'****************End of Function to Get DN of group****************