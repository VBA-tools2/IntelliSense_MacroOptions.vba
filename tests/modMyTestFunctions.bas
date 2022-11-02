Attribute VB_Name = "modMyTestFunctions"

Option Explicit

Public Function MyVBAFunction( _
    ByVal FirstArg As String, _
    ByVal AnotherArg As String _
        ) As String
Attribute MyVBAFunction.VB_Description = "A function described in XML"
Attribute MyVBAFunction.VB_ProcData.VB_Invoke_Func = " \n14"
    MyVBAFunction = "MyVBAFunction"
End Function

Public Function AnotherFunction( _
    ByVal FirstArg As String, _
    ByVal AnotherArg As String _
        ) As String
Attribute AnotherFunction.VB_Description = "A function described in XML"
Attribute AnotherFunction.VB_ProcData.VB_Invoke_Func = " \n14"
    AnotherFunction = "AnotherFunction"
End Function
