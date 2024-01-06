Attribute VB_Name = "Módulo1"
Sub macroPruebas()

    Dim diasGuardar As Integer, mesGuardar As Integer, yearGuardar As Integer
    Dim carpetaEntrada As String, carpetaSalida As String, datosEmpleados As String
    Dim archivosDatosEmpleados As String

    diasGuardar = Day(Date)
    mesGuardar = Month(Date)
    yearGuardar = Year(Date)
    
    carpetaEntrada = ThisWorkbook.Sheets("Main").Range("C3").Value
    carpetaSalida = ThisWorkbook.Sheets("Main").Range("C4").Value
    
    If carpetaEntrada = "" And carpetaSalida = "" Then
        MsgBox "Las carpetas de entrada y salida deben estar especificadas.", vbExclamation
        Exit Sub
    ElseIf Right(carpetaEntrada, 1) <> "\" Then
        carpetaEntrada = carpetaEntrada & "\"
        carpetaSalida = carpetaSalida & "\"
    End If
    
    datosEmpleados = carpetaEntrada & "Datos Empleados\"
    archivosDatosEmpleados = Dir(datosEmpleados & "*.*")
    
    Do While Len(archivosDatosEmpleados) > 0
    
        Application.DisplayAlerts = False
        Workbooks.OpenText Filename:=datosEmpleados & archivosDatosEmpleados
        Application.DisplayAlerts = True
        
        
        Windows(archivosDatosEmpleados).Activate
        ActiveWorkbook.Close SaveChanges:=False
        
        archivosDatosEmpleados = Dir()
    
    Loop

End Sub

