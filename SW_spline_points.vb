Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swSketchMgr As SketchManager
Dim filePath As String
Dim line As String
Dim fileNumber As Integer
Dim points As Variant

Sub main()

    ' Set the active SolidWorks application and model
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Ensure there is an open part document
    If swModel Is Nothing Or swModel.GetType <> swDocPART Then
        MsgBox "Please open a part document to run this macro."
        Exit Sub
    End If

    ' Set the path to the CSV file
    filePath = "FILE\PATH\HERE"  ' Update this path to the actual file location

    ' Read and parse the CSV file
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber

    ' Prepare an array to store points
    Dim pointArray() As Double
    Dim pointCount As Integer
    pointCount = 0

    ' Read each line and extract the coordinates
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        ' Skip header row
        If InStr(line, "Bone Name") = 0 Then
            Dim splitLine() As String
            splitLine = Split(line, ",")
            Dim x As Double, y As Double, z As Double
            x = CDbl(splitLine(1))
            y = CDbl(splitLine(2))
            z = CDbl(splitLine(3))
            
            ' Add the coordinates to the point array
            ReDim Preserve pointArray(pointCount * 3 + 2)
            pointArray(pointCount * 3) = x
            pointArray(pointCount * 3 + 1) = y
            pointArray(pointCount * 3 + 2) = z
            pointCount = pointCount + 1
        End If
    Loop

    ' Close the CSV file
    Close #fileNumber

    ' Begin a new 3D sketch
    Set swSketchMgr = swModel.SketchManager
    swSketchMgr.Insert3DSketch True

    ' Create a spline from the points
    If pointCount > 1 Then
        swSketchMgr.CreateSpline pointArray
    Else
        MsgBox "Not enough points to create a curve."
    End If

    ' End the 3D sketch
    swSketchMgr.Insert3DSketch True

    MsgBox "Curve created from CSV points."

End Sub