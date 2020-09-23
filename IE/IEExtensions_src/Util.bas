Attribute VB_Name = "Util"

Option Explicit



Public Function ToString() As String
    MsgBox "[" & Err.Number & "] " & Err.Description
    ToString = "[" & Err.Number & "] " & Err.Description
End Function
