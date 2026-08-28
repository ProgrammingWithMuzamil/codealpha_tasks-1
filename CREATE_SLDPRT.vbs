Option Explicit
Dim swApp, model, errors, warnings, ok, fso, folder, stepPath, outPath
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
stepPath = fso.BuildPath(folder, "bearing_housing.step")
outPath = fso.BuildPath(folder, "bearing_housing.SLDPRT")

If Not fso.FileExists(stepPath) Then
  MsgBox "bearing_housing.step was not found in this folder.", 16, "SolidWorks Converter"
  WScript.Quit 1
End If

On Error Resume Next
Set swApp = CreateObject("SldWorks.Application")
If Err.Number <> 0 Or swApp Is Nothing Then
  MsgBox "SolidWorks is not installed or could not be started.", 16, "SolidWorks Converter"
  WScript.Quit 2
End If
On Error GoTo 0

swApp.Visible = True
errors = 0
warnings = 0
On Error Resume Next
Set model = swApp.OpenDoc6(stepPath, 1, 0, "", errors, warnings)
If model Is Nothing Then
  Err.Clear
  errors = 0
  Set model = swApp.LoadFile4(stepPath, "", Nothing, errors)
End If
On Error GoTo 0

If model Is Nothing Then
  MsgBox "SolidWorks could not import the STEP file. Open bearing_housing.step manually and use Save As > SolidWorks Part (*.SLDPRT).", 16, "SolidWorks Converter"
  WScript.Quit 3
End If

On Error Resume Next
ok = model.SaveAs3(outPath, 0, 0)
If Err.Number <> 0 Then
  Err.Clear
  ok = False
End If
On Error GoTo 0

If ok Then
  MsgBox "Done. Created:" & vbCrLf & outPath, 64, "SolidWorks Converter"
Else
  MsgBox "The STEP file is open in SolidWorks. Use File > Save As and select SolidWorks Part (*.SLDPRT).", 48, "SolidWorks Converter"
End If
