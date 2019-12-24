Attribute VB_Name = "azs_Zip"
Option Explicit



Function Zip(myFileSpec, myZip)
' This function uses X-standards.com's X-zip component to add
' files to a ZIP file.
' If the ZIP file doesn't exist, it will be created on-the-fly.
' Compression level is set to maximum, only relative paths are
' stored.
'
' Arguments:
' myFileSpec    [string] the file(s) to be added, wildcards allowed
'                        (*.* will include subdirectories, thus
'                        making the function recursive)
' myZip         [string] the fully qualified path to the ZIP file
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
'
' The X-zip component is available at:
' http://www.xstandard.com/en/documentation/xzip/
' For more information on available functionality read:
' http://www.xstandard.com/printer-friendly.asp?id=C9891D8A-5390-44ED-BC60-2267ED6763A7

    Dim objZIP
    On Error Resume Next
    Err.Clear
    Set objZIP = CreateObject("XStandard.Zip")
    objZIP.Pack myFileSpec, myZip, , , 9
    Zip = Err.Number
    Err.Clear
    Set objZIP = Nothing
    On Error GoTo 0
End Function

Function UnZip(myFileSpec, myZip)
' This function uses X-standards.com's X-zip component to add
' files to a ZIP file.
' If the ZIP file doesn't exist, it will be created on-the-fly.
' Compression level is set to maximum, only relative paths are
' stored.
'
' Arguments:
' myFileSpec    [string] the file(s) to be added, wildcards allowed
'                        (*.* will include subdirectories, thus
'                        making the function recursive)
' myZip         [string] the fully qualified path to the ZIP file
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
'
' The X-zip component is available at:
' http://www.xstandard.com/en/documentation/xzip/
' For more information on available functionality read:
' http://www.xstandard.com/printer-friendly.asp?id=C9891D8A-5390-44ED-BC60-2267ED6763A7

    Dim objZIP
    On Error Resume Next
    Err.Clear
    Set objZIP = CreateObject("XStandard.Zip")
    objZIP.UnPack myFileSpec, myZip, "*.*"
    
'    objZIP.UnPack myFileSpec, myZip, , , 9
    UnZip = Err.Number
    Err.Clear
    Set objZIP = Nothing
    On Error GoTo 0
End Function
