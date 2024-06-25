' Este script fue desarrollado cuando era pasante de ingeniería en Peasa.
' Compartí este código para automatizar archivos con Visual Basic for Applications (VBA).
' Esta solución fue eficiente y nos ayudó en el Departamento de Producción.


Dim TiempoInicio As Double
Dim FilaActual As Integer
Dim ColumnaActual As Integer
Dim Hoja As Worksheet
Dim TemporizadorCiclico As Boolean
Dim TemporizadorActivo As Boolean
Dim FilaStart As Integer
Dim Reset As Boolean

Sub IniciarTemporizadorCICLICO()
    ' Poner el valor inicial del temporizador (00:00:00) en la celda B7
    Set Hoja = ActiveSheet
    Hoja.Range("D8").Value = Format(TimeValue("00:00:00"), "hh:mm:ss")
    ' Guardar el tiempo actual cuando se inicia el temporizador
    TiempoInicio = Now
    ' Inicializar la fila y columna actual
    FilaActual = 10
    ColumnaActual = 4
    ' Activar el temporizador
    TemporizadorCiclico = True
    ' Iniciar el temporizador llamando a la función ActualizarCeldaB7 cada segundo
    'Application.OnTime Now + TimeValue("00:00:01"), "ActualizarCeldaB7"
End Sub
Sub CambioDeActividadCiclico()
    ' Verificar si el temporizador está activo
    If Not TemporizadorCiclico Then Exit Sub
    ' Encontrar el último valor en la columna C hasta la fila 30
    Dim UltimaFila As Long
    UltimaFila = Hoja.Cells(30, 3).End(xlUp).Row
    ' Verificar si la columna C tiene algún valor
    If UltimaFila >= FilaActual Then
        ' Actualizar la celda actual con el tiempo transcurrido
        Hoja.Cells(FilaActual, ColumnaActual).Value = Format(Now - TiempoInicio, "hh:mm:ss")
        FilaActual = FilaActual + 2
        ' Verificar si la fila actual supera la última fila con datos en la columna C
        If FilaActual > UltimaFila Then
            ' Pasar a la siguiente columna y reiniciar la fila
            ColumnaActual = ColumnaActual + 1
            FilaActual = 8
        End If
     End If
End Sub

Sub DetenerTemporizadorCiclico()
    ' Cancelar el temporizador
    On Error Resume Next
    'Application.OnTime Now + TimeValue("00:00:01"), "ActualizarCeldaCiclico", , False
    On Error GoTo 0
    ' Desactivar el temporizador
    
    Reset = TemporizadorCiclico
    If Reset = True Then
        TiempoFinal = Now
        Hoja.Range("T3").Value = Format(TiempoFinal - TiempoInicio, "hh:mm:ss")
        TemporizadorCiclico = False
        
    End If
End Sub
Sub ObtenerTiempos()
    On Error GoTo ErrorHandler
    Dim FilaResultado As Integer
    Dim TiempoInicioFila As Date
    Dim TiempoFinalFila As Date
    Dim ColumnaActual As Integer
    Dim Hoja As Worksheet

    ' Establecer la hoja activa
    Set Hoja = ActiveSheet

    ' Inicializar la fila para los resultados
    FilaResultado = 9

    ' Iterar sobre las columnas comenzando desde la columna D
    For ColumnaActual = 4 To Hoja.Columns.Count
        ' Repetir el proceso mientras haya datos en la columna actual
        Do While Hoja.Cells(FilaResultado + 1, ColumnaActual).Value <> ""
            ' Obtener el tiempo de inicio y final de la fila actual
            TiempoInicioFila = Hoja.Cells(FilaResultado - 1, ColumnaActual).Value
            TiempoFinalFila = Hoja.Cells(FilaResultado + 1, ColumnaActual).Value

            ' Calcular la diferencia de tiempo
            Dim Diferencia As Double
            Diferencia = TiempoFinalFila - TiempoInicioFila

            ' Colocar el resultado en la celda especificada con el formato adecuado
            Hoja.Cells(FilaResultado, ColumnaActual).Value = Format(Diferencia, "hh:mm:ss")

            ' Pasar a la siguiente fila
            FilaResultado = FilaResultado + 2
        Loop

        ' Si hay datos en la celda de la derecha, pasar a la siguiente columna
        If Hoja.Cells(8, ColumnaActual + 1).Value <> "" Then
            ' Obtener el tiempo de la última fila con datos en la columna C
            TiempoInicioFila = Hoja.Cells(Hoja.Cells(30, 3).End(xlUp).Row, ColumnaActual).Value
            
            ' Obtener el tiempo de la primera fila de la siguiente columna (fila 8)
            TiempoFinalFila = Hoja.Cells(8, ColumnaActual + 1).Value

            ' Calcular la diferencia de tiempo
            Dim DiferenciaSiguienteColumna As Double
            DiferenciaSiguienteColumna = TiempoFinalFila - TiempoInicioFila

            ' Colocar el resultado en la celda especificada con el formato adecuado
            Hoja.Cells(Hoja.Cells(30, 3).End(xlUp).Row + 1, ColumnaActual).Value = Format(DiferenciaSiguienteColumna, "hh:mm:ss")
        End If

        ' Reiniciar la fila para la siguiente columna
        FilaResultado = 9
    Next ColumnaActual

    Exit Sub ' Salir del procedimiento normalmente

ErrorHandler:
    MsgBox "termino de la obtencion de tiempos"
    On Error Resume Next
End Sub
Sub EliminaTiempos()

Range("D8:BM31").Select
    Selection.ClearContents

    Range("D33:BM33").Select
    Selection.ClearContents
    
End Sub
Sub EliminaDatosTabla()
' EliminaDatosTabla Macro

    Range("D8:BM31").Select
    Selection.ClearContents
    
    Range("D33:BM33").Select
    Selection.ClearContents
    
End Sub


