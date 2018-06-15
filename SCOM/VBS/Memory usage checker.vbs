Const adOpenStatic = 3

Const adLockOptimistic = 3
Set oAPI = CreateObject(“MOM.ScriptAPI”)
Set oBag = oAPI.CreatePropertyBag()
Set objConnection = CreateObject(“ADODB.Connection”)
Set objRecordSet = CreateObject(“ADODB.Recordset”)
objConnection.Open _
“Provider=SQLOLEDB;Data Source=mobile-opsmgr;” & _
“Trusted_Connection=Yes;Initial Catalog=TechDaysdb;” & _
“User ID=domain\username;Password=password;”
objRecordSet.Open “SELECT * FROM OpsMgrV”, _
objConnection, adOpenStatic, adLockOptimistic
varNo = objRecordSet.RecordCount
If varNO > 4 Then
Rows = varNo
Call oBag.AddValue(“Status”,”Error”)
Call oBag.AddValue(“Rows”,Rows)
Call oAPI.Return(oBag)
Else
Call oBag.AddValue(“Status”,”Ok”)
Call oAPI.Return(oBag)
End If