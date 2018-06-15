Option Explicit  

Dim oAPI, oBag
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()
 
Dim adoConn, adoRS, adoCmd, connStr, strSql, strResult, iLength
connStr = "Provider=SQLOLEDB.1; Data Source=Ahsl250;Initial Catalog=ASPIRE;User ID=egateprod;Password=c0ntr0l;"
strSQL = "SET NOCOUNT ON DECLARE @FailedItemsTable TABLE (ItemID INT, Status VARCHAR(50)) DECLARE @ItemID INT DECLARE @Status VARCHAR(50) DECLARE @AggregatedRows VARCHAR(MAX) SET @AggregatedRows = '' INSERT INTO @FailedItemsTable (ItemID, Status) SELECT i.ItemID AS 'ItemID', st.Name AS 'Status' FROM Item i INNER JOIN ItemType it ON it.ItemTypeID = i.ItemTypeID INNER JOIN System s ON s.SystemID = it.SystemID INNER JOIN ItemStatusHistory ish ON ish.ItemID = i.ItemID INNER JOIN Status st ON st.StatusID = ish.StatusID WHERE s.Code IN ('RESPA') AND dbo.uf_GetItemStatusCode(i.ItemID) IN (SELECT s.Code FROM Status s INNER JOIN ViewModeStatus vms ON s.StatusID = vms.StatusID INNER JOIN ViewMode vm ON vm.ViewModeID = vms.ViewModeID WHERE vm.Code IN ('TECHNICAL', 'FUNCTIONAL')) AND ish.ItemStatusHistoryID IN (SELECT MAX(ish2.ItemStatusHistoryID) FROM ItemStatusHistory ish2 WHERE ish2.ItemID = i.ItemID) DECLARE ItemCursor CURSOR FOR SELECT ItemID, Status FROM @FailedItemsTable ORDER BY ItemID OPEN ItemCursor FETCH NEXT FROM ItemCursor INTO @ItemID, @Status WHILE @@FETCH_STATUS = 0 BEGIN SET @AggregatedRows = @AggregatedRows + '""' + CONVERT(VARCHAR, @ItemID) + '"",""' + @Status + '""' + CHAR(13) + CHAR(10) FETCH NEXT FROM ItemCursor INTO @ItemID, @Status END CLOSE ItemCursor DEALLOCATE ItemCursor SELECT LEN(@AggregatedRows) AS StringLength, @AggregatedRows As Result"

 

Set adoConn = CreateObject("ADODB.Connection")
adoConn.ConnectionString = connStr
adoConn.Open

Set adoCmd = CreateObject("ADODB.Command")
Set adoCmd.ActiveConnection = adoConn
adoCmd.CommandText = strSql 
Set adoRS = adoCmd.Execute 

adoRS.MoveFirst
 
iLength = CInt(adoRS.Fields("StringLength").Value)
strResult = adoRS.Fields("Result").Value

If iLength &gt; 0 Then
                Call oBag.AddValue("State", "BAD")
Else
                Call oBag.AddValue("State", "OK")
End If
 

Call oBag.AddValue("Result", strResult) 

adoRS.Close
adoConn.Close 

Call oAPI.Return(oBag)