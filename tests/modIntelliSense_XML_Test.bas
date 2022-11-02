Attribute VB_Name = "modIntelliSense_XML_Test"

Option Explicit
Option Private Module

'@TestModule
'@Folder("IntelliSense.Tests")

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

Private XmlDirectory As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
#If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
#Else
        Set Assert = New Rubberduck.PermissiveAssertClass
        Set Fakes = New Rubberduck.FakesProvider
#End If
    XmlDirectory = ThisWorkbook.Path & "\XMLs\"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'REF: <https://stackoverflow.com/a/28237845/5776000>
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
Private Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(FileName) And vbDirectory) <> vbDirectory)
    On Error GoTo 0
End Function

'==============================================================================
'@TestMethod("Invalid XML File")
Private Sub XmlFileNotValid_XmlFileNotAString_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrNotAnXmlFile
    On Error GoTo TestFail
    
    RegisterFunctionsFromXmlFile 123

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Invalid XML File")
Private Sub XmlFileNotPresent_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrXmlFileDoesntExist
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = "c:\I_am_not_there.xml"
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Parser Error")
Private Sub XmlFileNotValid_ClosingIntelliSenseTagTypo_RaiseError()
    Const ExpectedError As Long = -1072896659
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "ParserError_ClosingIntelliSenseTagTypo.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Parser Error")
Private Sub XmlFileNotValid_ClosingFunctionInfoTagMissing_RaiseError()
    Const ExpectedError As Long = -1072896659
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "ParserError_ClosingFunctionInfoTagMissing.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DTD Error")
Private Sub XmlFileNotValid_DtdNotAllowed_RaiseError()
    Const ExpectedError As Long = -1072896636
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlError_DtdError.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Schema Error")
Private Sub XsdFileNotValid_NoSchema_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrNoOrWrongSchema
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_NoSchema.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Schema Error")
Private Sub XsdFileNotValid_WrongSchema_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrNoOrWrongSchema
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_WrongSchema.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XSD Error")
Private Sub XmlFileNotValid_FunctionNameMissing_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrNoFunctionName
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_FunctionNameMissing.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XSD Error")
Private Sub XmlFileNotValid_FunctionDescriptionMissing_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrNoFunctionDescription
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_FunctionDescriptionMissing.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XSD Error")
Private Sub XmlFileNotValid_FunctionDescriptionTooLong_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrStringTooLong
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_FunctionDescriptionTooLong.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XML Error")
Private Sub XmlFileNotValid_FunctionDoesntExist_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrFunctionDoesntExist
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlError_FunctionDoesntExist.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XML Error")
Private Sub XmlFileNotValid_CategoryNumberTooLow_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrInvalidCategoryNumber
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlError_CategoryNumberTooLow.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XML Error")
Private Sub XmlFileNotValid_CategoryNumberTooHigh_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrInvalidCategoryNumber
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlError_CategoryNumberTooHigh.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("XSD Error")
Private Sub XsdFileNotValid_CategoryNameTooLong_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrStringTooLong
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_CategoryNameTooLong.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'TODO: add test for `getFunctionHelpTopic`

'@TestMethod("XSD Error")
Private Sub XsdFileNotValid_ArgumentDescriptionTooLong_RaiseError()
    Const ExpectedError As Long = eIntelliSenseError.ErrStringTooLong
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XsdError_ArgumentDescriptionTooLong.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Works fine")
Private Sub XmlFileValid_OneArgumentTooLess_WorksFine()                        'TODO Rename test
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlFile_OneArgumentTooLess.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Works fine")
Private Sub XmlFileValid_OneArgumentTooMuch_WorksFine()                        'TODO Rename test
    On Error GoTo TestFail
    
    Dim XmlFile As String
    XmlFile = XmlDirectory & "XmlFile_OneArgumentTooMuch.xml"
    If Not FileExists(XmlFile) Then Assert.Inconclusive
    
    RegisterFunctionsFromXmlFile XmlFile
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

