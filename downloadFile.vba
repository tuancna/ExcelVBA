Public Function getLink(r As Range)
    getLink = r.Hyperlinks(1).Address
End Function

Public Function downloadFile(url As String, destination As String, fileName As String) As String
    'Check destination
    If Dir(destination, vbDirectory) = Empty Then
        On Error GoTo pathErr
        MkDir destination
    End If
    
    'Download
    Path = destination & fileName
    
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    Set oBinaryStream = CreateObject("ADODB.Stream")
    adTypeBinary = 1
    oBinaryStream.Type = adTypeBinary
    
    On Error GoTo httpErr
    oXMLHTTP.Open "GET", url, False
    oXMLHTTP.Send
    aBytes = oXMLHTTP.responsebody
    
    On Error GoTo 0
    oBinaryStream.Open
    oBinaryStream.Write aBytes
    adSaveCreateOverWrite = 2
    oBinaryStream.SaveToFile Path, adSaveCreateOverWrite
    oBinaryStream.Close
    
    'Notify
    downloadFile = "Downloaded."
    Exit Function
    
    'Error
pathErr:
    downloadFile = "Unable to create new folder."
    Exit Function
    
httpErr:
    downloadFile = "Unable to download."
    Exit Function
End Function
