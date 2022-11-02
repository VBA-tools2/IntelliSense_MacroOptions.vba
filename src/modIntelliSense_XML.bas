Attribute VB_Name = "modIntelliSense_XML"

'@Folder("IntelliSense")

Option Explicit

'add references
'- Microsoft Scripting Runtime
'- Microsoft XML, 6.0

'have a look at <https://stackoverflow.com/a/53559474>
'and at <https://stackoverflow.com/a/51919182>

'maybe relevant for namespaces issues
'<https://www.vbforums.com/showthread.php?678257-RESOLVED-XML-ReplaceChild-automatically-adding-unwanted-namespace-attribute&s=298800ee42a6c9e209b0a128482f03ec&p=4165101&viewfull=1#post4165101>

'other interesting information on XML files
'<https://docs.microsoft.com/en-us/previous-versions/troubleshoot/msxml/disable-output-escaping-style-sheet>

'=====
'links to hopefully make XInclude work (at least in VBA). The hint was given in
'   <https://stackoverflow.com/a/14694298>
'- <https://bettersolutions.com/excel/xml/xml-transforms-xslt.htm>
'- <https://social.msdn.microsoft.com/Forums/de-DE/ba46717f-66ee-46fd-a7eb-df4de17d9371/excel-vba-coding-for-xls-transformation?forum=isvvba>

'=====
'- get category from comments in XML file --> done and works
'- same for `HelpContextID`?

'- check what for IntelliSense has precedence: XML or '__IntelliSense__' sheet

'namespace prefix to make XPath searches work
Private Const NameSpacePrefix As String = "doc:"

Public Enum eIntelliSenseError
    [_First] = vbObjectError + 1
    ErrNotAnXmlFile = [_First]
    ErrXmlFileDoesntExist
    ErrNoOrWrongSchema
    ErrNoFunctionName
    ErrNoFunctionDescription
    ErrStringTooLong
    ErrFunctionDoesntExist
    ErrInvalidCategoryNumber
    [_Last] = ErrInvalidCategoryNumber
End Enum

'BUG: delete me when the Unit Tests are written
Public Sub bla()
    Dim XmlDirectory As String
    XmlDirectory = ThisWorkbook.Path & "\XMLs\"
    RegisterFunctionsFromXmlFile XmlDirectory & "MyTest.IntelliSense.xml"
End Sub

'from <https://stackoverflow.com/a/5747032>
Public Sub RegisterFunctionsFromXmlFile(Optional ByVal XmlFile As String = vbNullString)
    
    If Len(XmlFile) = 0 Then
        Dim xDocName As String
        xDocName = GetIntellisenseFileName
    ElseIf Not IsAnXmlFileString(XmlFile) Then
        RaiseErrorNotAnXmlFile XmlFile
    ElseIf Not FileExists(XmlFile) Then
        RaiseErrorXmlFileDoesntExist XmlFile
    Else
        xDocName = XmlFile
    End If
    
    'see <http://exceldevelopmentplatform.blogspot.com/2017/12/fake-namespace-vba-msxml2-xpath.html>
    'plus the links given there
    Dim xNamespaces As String
    xNamespaces = _
            "xmlns:" & _
            Left$(NameSpacePrefix, Len(NameSpacePrefix) - 1) & _
            "='http://schemas.excel-dna.net/intellisense/1.0'"
    
    Dim xDoc As MSXML2.DOMDocument60
    Set xDoc = New MSXML2.DOMDocument60
    
    With xDoc
        .validateOnParse = True
        .setProperty "SelectionNamespaces", xNamespaces
        .async = False
        
        If Not .Load(xDocName) Then GoTo errHandler
    End With
    
    GetFunctions xDoc

tidyUp:
    Exit Sub
    
errHandler:
    Err.Raise xDoc.parseError.ErrorCode, , xDoc.parseError.reason
    
End Sub

Private Function IsAnXmlFileString(ByVal XmlFile As String) As Boolean
    IsAnXmlFileString = (Right$(XmlFile, 4) = ".xml")
End Function

'REF: <https://stackoverflow.com/a/28237845/5776000>
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
Private Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(FileName) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function

Private Function GetIntellisenseFileName() As String

    Dim fso As New Scripting.FileSystemObject
    With ThisWorkbook
        GetIntellisenseFileName = _
                .Path & "\" & _
                fso.GetBaseName(.Name) & _
                ".IntelliSense.xml"
    End With
    
End Function

'NOTE: not used here. Could be an alternative to to the `Err.Raise` in `LoadDocument`
'modified from <https://stackoverflow.com/a/53559474>
Private Sub ShowXmlParsingError(ByVal xDoc As MSXML2.DOMDocument60)
    Dim xPE As MSXML2.IXMLDOMParseError
    Set xPE = xDoc.parseError
    With xPE
        Dim strErrText As String
        strErrText = "Load error " & .ErrorCode & " xml file " & vbCrLf & _
                Replace(.Url, "file:///", vbNullString) & vbCrLf & vbCrLf & _
                .reason & _
                "Source Text: " & .srcText & vbCrLf & vbCrLf & _
                "Line No.:    " & .Line & vbCrLf & _
                "Line Pos.: " & .linepos & vbCrLf & _
                "File Pos.:  " & .filepos & vbCrLf & vbCrLf
    End With
    Set xPE = Nothing
    MsgBox strErrText, vbExclamation
End Sub

Private Sub GetFunctions(ByVal xDoc As MSXML2.DOMDocument60)

    '==========================================================================
    Const FunctionString As String = "//" & NameSpacePrefix & "Function"
    '==========================================================================
    
    Dim xDocName As String
    xDocName = Right$(xDoc.Url, Len(xDoc.Url) - InStrRev(xDoc.Url, "/"))
    
    Dim root As MSXML2.IXMLDOMElement
    Set root = xDoc.DocumentElement
    
    Dim ElementList As MSXML2.IXMLDOMNodeList
    Set ElementList = root.SelectNodes(FunctionString)
    
    If ElementList.Length = 0 Then RaiseErrorNoOrWrongSchema xDoc.Url
    
    Dim Element As MSXML2.IXMLDOMNode
    For Each Element In ElementList
        AddUDF Element, xDocName
    Next

End Sub

Private Sub AddUDF( _
    ByVal Element As MSXML2.IXMLDOMElement, _
    ByVal xDocName As String _
)

    Dim FunctionName As String
    FunctionName = getFunctionName(Element, xDocName)
    
    Dim FunctionDescription As String
    FunctionDescription = getFunctionDescription(Element, xDocName)
    
    Dim FunctionCategory As Variant
    FunctionCategory = getFunctionCategory(Element, xDocName)
    
    Dim arrArgumentDescriptions As Variant
    arrArgumentDescriptions = getArgumentDescriptions(Element, xDocName)
    
    On Error GoTo errHandler
    If UBound(arrArgumentDescriptions) >= 0 Then
        Application.MacroOptions _
                Category:=FunctionCategory, _
                Macro:=FunctionName, _
                Description:=FunctionDescription, _
                ArgumentDescriptions:=arrArgumentDescriptions
    Else
        Application.MacroOptions _
                Category:=FunctionCategory, _
                Macro:=FunctionName, _
                Description:=FunctionDescription
    End If
    On Error GoTo 0
    
tidyUp:
    Exit Sub
    
errHandler:
    If Err.Number = 1004 Then
        RaiseErrorFunctionDoesntExist xDocName, FunctionName
    Else
        Err.Raise Number:=Err.Number, Description:=Err.Description
    End If

End Sub

Private Function getFunctionName( _
    ByVal Element As MSXML2.IXMLDOMNode, _
    ByVal xDocName As String _
        ) As String
    
    On Error GoTo errHandler
    Dim result As String
    result = Element.Attributes.getNamedItem("Name").Text
    On Error GoTo 0
    
    getFunctionName = result
    
tidyUp:
    Exit Function
    
errHandler:
    RaiseErrorNoFunctionName xDocName
    
End Function

Private Function getFunctionDescription( _
    ByVal Element As MSXML2.IXMLDOMNode, _
    ByVal xDocName As String _
        ) As String
    
    On Error GoTo errHandler
    Dim result As String
    result = Element.Attributes.getNamedItem("Description").Text
    On Error GoTo 0
    
    If Not IsLengthOk(result) Then RaiseErrorStringTooLong xDocName
    
    getFunctionDescription = result
    
tidyUp:
    Exit Function
    
errHandler:
    RaiseErrorNoFunctionDescription xDocName
    
End Function

Private Function getFunctionCategory( _
    ByVal Element As MSXML2.IXMLDOMNode, _
    ByVal xDocName As String _
        ) As Variant
    
    '==========================================================================
    Const AttributeName As String = "Category"
    '==========================================================================
    
    Dim testAttribute As MSXML2.IXMLDOMAttribute
    Set testAttribute = Element.Attributes.getNamedItem("Category")
    
    If testAttribute Is Nothing Then
        'put in default (user defined) category
        getFunctionCategory = 14
    Else
        Dim result As String
        result = testAttribute.Text
    
        If IsInteger(result) Then
            If Int(result) < 1 Or Int(result) > 32 Then RaiseErrorInvalidCategoryNumber xDocName, result
        Else
            If Not IsLengthOk(result) Then RaiseErrorStringTooLong xDocName
        End If
        
        getFunctionCategory = result
    End If
        
End Function

Private Function IsInteger( _
    ByVal testCategory As String _
        ) As Boolean
    
    IsInteger = False
    
    If Not IsNumeric(testCategory) Then Exit Function
    If testCategory <> Int(testCategory) Then Exit Function
    
    IsInteger = True
    
End Function

Private Function getFunctionHelpTopic( _
    ByVal Element As MSXML2.IXMLDOMNode, _
    ByVal xDocName As String _
        ) As String

    '==========================================================================
    Const AttributeName As String = "HelpTopic"
    '==========================================================================

    Dim testAttribute As MSXML2.IXMLDOMAttribute
    Set testAttribute = Element.Attributes.getNamedItem(AttributeName)
    
    If testAttribute Is Nothing Then Exit Function

    Dim result As String
    result = testAttribute.Text
    
    If Not IsLengthOk(result) Then RaiseErrorStringTooLong xDocName
    
    getFunctionHelpTopic = result

End Function

'NOTE: not needed for this purpose, but could be used for checking the XML file
'      (if argument names are present and maybe checking their length)
Private Function getArgumentNames( _
    ByVal Element As MSXML2.IXMLDOMNode _
        ) As Variant

    If Element.HasChildNodes Then
        Dim NoOfArguments As Long
        NoOfArguments = Element.ChildNodes.Length
        
        Dim arrNames() As Variant
        ReDim arrNames(0 To NoOfArguments - 1)
        
        Dim i As Long
        For i = LBound(arrNames) To UBound(arrNames)
            With Element.ChildNodes.Item(i).Attributes
                arrNames(i) = .getNamedItem("Name").Text
            End With
        Next
        
        getArgumentNames = arrNames
    Else
        getArgumentNames = Array()
    End If

End Function

Private Function getArgumentDescriptions( _
    ByVal Element As MSXML2.IXMLDOMNode, _
    ByVal xDocName As String _
        ) As Variant

    If Element.HasChildNodes Then
        Dim NoOfArguments As Long
        NoOfArguments = Element.ChildNodes.Length
        
        Dim arrDescriptions() As Variant
        ReDim arrDescriptions(0 To NoOfArguments - 1)
        
        Dim i As Long
        For i = LBound(arrDescriptions) To UBound(arrDescriptions)
            With Element.ChildNodes.Item(i).Attributes
                Dim result As String
                result = .getNamedItem("Description").Text
                
                If Not IsLengthOk(result) Then RaiseErrorStringTooLong xDocName
                
                arrDescriptions(i) = result
            End With
        Next
        
        getArgumentDescriptions = arrDescriptions
    Else
        getArgumentDescriptions = Array()
    End If

End Function

Private Function IsLengthOk( _
    ByVal ToCheckString As String _
        ) As Boolean
    IsLengthOk = (Len(ToCheckString) <= 255)
End Function

'==============================================================================
Private Sub RaiseErrorNotAnXmlFile(ByVal XmlFile As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrNotAnXmlFile, _
            Description:= _
                    "The file '" & XmlFile & "' must end with '.xml'."
End Sub

Private Sub RaiseErrorXmlFileDoesntExist(ByVal XmlFile As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrXmlFileDoesntExist, _
            Description:= _
                    "The file '" & XmlFile & "' doesn't exist."
End Sub

Private Sub RaiseErrorNoOrWrongSchema(ByVal Url As String)
    Dim XmlFile As String
    XmlFile = Right$(Url, Len(Url) - InStrRev(Url, "/"))
    
    Err.Raise _
            Number:=eIntelliSenseError.ErrNoOrWrongSchema, _
            Description:= _
                    "The file '" & XmlFile & "' has no elements in the expected namespace." & _
                    vbCrLf & _
                    "Maybe there is no namespace or the given one is wrong?"
End Sub

Private Sub RaiseErrorNoFunctionName(ByVal xDocName As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrNoFunctionName, _
            Description:= _
                    "One Function in '" & xDocName & "' has no 'Name' attribute." & _
                    vbCrLf & _
                    "Please check it against the the corresponding XSD file."
End Sub

Private Sub RaiseErrorNoFunctionDescription(ByVal xDocName As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrNoFunctionDescription, _
            Description:= _
                    "One Function in '" & xDocName & "' has no 'Description' attribute." & _
                    vbCrLf & _
                    "Please check it against the the corresponding XSD file."
End Sub

Private Sub RaiseErrorStringTooLong(ByVal xDocName As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrStringTooLong, _
            Description:= _
                    "One string in '" & xDocName & "' is too long (i.e. longer than 255 chars)." & _
                    vbCrLf & _
                    "Please check it against the the corresponding XSD file."
End Sub

Private Sub RaiseErrorFunctionDoesntExist(ByVal xDocName As String, ByVal FunctionName As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrFunctionDoesntExist, _
            Description:= _
                    "Most likely the function '" & FunctionName & "' from/in '" & xDocName & "' doesn't exist." & _
                    vbCrLf & _
                    "Is it a typo?"
End Sub

Private Sub RaiseErrorInvalidCategoryNumber(ByVal xDocName As String, ByVal CategoryNumber As String)
    Err.Raise _
            Number:=eIntelliSenseError.ErrInvalidCategoryNumber, _
            Description:= _
                    "The category number '" & CategoryNumber & "' in file '" & xDocName & "' is invalid." & _
                    vbCrLf & _
                    "Only numbers from 1 to 32 are allowed."
End Sub
