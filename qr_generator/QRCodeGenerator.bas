Attribute VB_Name = "Module2"
Sub vietqr()
Attribute vietqr.VB_Description = "Macro t?o mã QR cho khách hàng"
Attribute vietqr.VB_ProcData.VB_Invoke_Func = " \n14"
'
' vietqr Main macro
' Create QR Code for customer
'
    Dim http As Object ' API variables
    Dim apiKey As String
    Dim jsonData As String
    Dim ResponseText As String
    Dim QRCodeDataURL As Dictionary
    Dim DataURL As String
    Dim payment As Double ' Customer data
    Dim payee_name As String
    Dim school As String
    Dim class As String
    Dim PayInfo As String
    Dim paycontent As String
    
    Dim ws As Worksheet ' Workbook variable
    Dim cell As Range
    Dim i As Long

    ' Set the workbook sheeet where data is
    Set ws = ThisWorkbook.Worksheets("RawData")
    ' Set the array of columns to iterate over
    For i = 2 To 4:
        payment = ws.Cells(2, "G").Value
        payee_name = ws.Cells(2, "D").Value
        school = ws.Cells(2, "B").Value
        class = ws.Cells(2, "C").Value
        PayInfo = ws.Cells(2, "F").Value
        
        paycontent = payee_name & "_" & school & "_" & class & "_" & PayInfo
        ' MsgBox paycontent -- Double-check value
        
        ' JSON content to be sent in the request body
        jsonData = "{""accountNo"": 113366668888, " & _
                    """accountName"": ""EMG Education"", " & _
                    """acqId"": 970415, " & _
                    """amount"": " & payment & ", " & _
                    """addInfo"": """ & paycontent & """, " & _
                    """template"": ""compact""}"
        ' MsgBox jsonData -- REMINDER: make sure to check if JSON receive the variable rather than hardcoding in the name of the variables
        ' Create a new HTTP request object, XMLHTTP object is suitable for most API call
        
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
            
        ' Send the request
        http.Open "POST", "https://api.vietqr.io/v2/generate", False
        http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "X-Api-Key", "4d927d20-3f73-4cf6-b5bd-e25f13409638" ' TODO: remember to set the api again to retrieve than hardcoding it
        http.setRequestHeader "x-client-id", "49ec4962-db36-4f25-ae84-455668577ca7"
        http.send jsonData
            
        ' -- TEST DRIVE FOR SUCCESSFUL JSON RESPONSE --
        ' MsgBox http.ResponseText ' --- TEST: PASSED ---
        
        ResponseText = http.ResponseText
            
        ' Parse the response to extract the QR code data URL
        Set QRCodeDataURL = JsonConverter.ParseJson(ResponseText)
            
        ' Slice to the data section of the JSON Response
        DataURL = QRCodeDataURL("data")("qrDataURL")
            
        ' -- TEST DRIVE FOR SUCCESFUL EXTRACTION OF JSON DATAURL to STRING --
        ' MsgBox DataURL ' --- TEST: PASSED ---
        Dim Directory As String
        Directory = "D:\emgpt\vba_qr\vba-qr_test_" & i - 1 & ".png"
        ' Save the QR code image
        SaveDataUrlAsImage DataURL, Directory
            
        ' Clean up
        Set http = Nothing
    Next i
End Sub

Sub SaveDataUrlAsImage(DataURL As String, filePath As String)
    ' This subroutine saves an image from a data URL to a file
    
    Dim imageData As String
    Dim b64Data() As Byte
    Dim objFSO As Object
    Dim objStream As Object
    
    ' Extract base64 QRdata from data URL by indexing the right part
    ' Here, I eyeballed the left redundant part and counted that as 22 to omit
    imageData = Right(DataURL, Len(DataURL) - 22)
    ' -- TEST DRIVE FOR CORRECT INDEXING --
    ' MsgBox imageData
    
    ' Convert base64 data to binary
    b64Data = DecodeBase64(imageData)
    ' MsgBox b64Data
    
    ' Convert the decoded byte code into image, .png format
    MyFile = filePath
    ' Get the next available file number
    FileNum = FreeFile

    ' Open the file in output mode ('Output' means writing to a file)
    Open MyFile For Binary Access Write As #FileNum
    ' Write the buffer (Byte Array) to the file
    Put #FileNum, 1, b64Data
    
    ' Close the file
    Close #FileNum
End Sub

'' For Reference on how to write the byte array into the file
' Function WriteByteArrToFile(filePath As String, buffer() As Byte) As Boolean
'    Dim fileNmb As Integer
'    On Error GoTo ErrorHandler
'
'    ' Get a free file number
'    fileNmb = FreeFile
'    ' Open the file for binary access (write mode)
'    Open filePath For Binary Access Write As #fileNmb
'    ' Write the buffer (Byte array) to the file
'    Put #fileNmb, 1, buffer
'    ' Close the file
'    Close #fileNmb
'
'    ' Return success
'    WriteByteArrToFile = True
'    Exit Function
'
'ErrorHandler:
'    ' Handle any errors (e.g., file cannot be accessed for writing)
'    WriteByteArrToFile = False
'    End Function
' --------------------------------------------------------------------------------

Function DecodeBase64(base64String As String) As Byte()
    Dim objXML As Object
    Dim objNode As Object
    
    ' Create XML Document object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    
    ' Create node for base64 string
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = base64String
    
    ' Return decoded byte array
    DecodeBase64 = objNode.NodeTypedValue
    
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Sub DownloadImageFromURL(URL As String, filePath As String)
    ' This sub downloads an image from a URL to a file
    
    Dim WinHttpReq As Object
    Dim oStream As Object
    
    ' Create HTTP request object
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Send HTTP request to download image
    WinHttpReq.Open "GET", URL, False
    WinHttpReq.send
    
    ' Create a stream object to write the binary data to a file
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 ' Binary
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile filePath, 2 ' Overwrite
    oStream.Close
End Sub

