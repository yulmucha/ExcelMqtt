Sub Publish_xxx_Button_Click()
Dim sheet As Worksheet
Dim tools As ExcelMqtt
Set sheet = ActiveSheet
Set tools = New ExcelMqtt

tools.SetExcelFileName (Application.ActiveWorkbook.FullName)
tools.SetBrokerHostName ("xxx")
tools.SetBrokerPort (1884)
tools.SetClientId ("ExcelVBA0001")
tools.SetUsername ("mqttuser9")
tools.SetPassword ("xxxx")
tools.SetTopic ("MDM/EXCEL/BUS-SHELTER/LOAD")
tools.SetStartRecordRow (3)
tools.SetPropertyRow (2)
tools.SetChunkSize (500)
tools.SetExcelSheetName ("GPS")

tools.Publish
End Sub