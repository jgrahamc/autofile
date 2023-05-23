Attribute VB_Name = "mdlDRainbow"
Public Declare Function Go Lib "drainbow.dll" (ByVal eclass As String) As Long
Public Declare Function GetMailFilename Lib "drainbow.dll" (ByVal mailfile As String, ByVal size As Long) As Long

