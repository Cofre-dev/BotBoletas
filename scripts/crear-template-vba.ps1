# ═══════════════════════════════════════════════════════════════════════════
#  crear-template-vba.ps1
#  Crea templates/vba_template.xlsm con las macros del Bot Boletas SII.
#  Requiere Microsoft Excel instalado.
#
#  IMPORTANTE: Si Excel lanza error de seguridad sobre VBProject, habilita:
#    Excel → Archivo → Opciones → Centro de confianza
#    → Configuración del centro de confianza → Configuración de macros
#    → [✓] Confiar en el acceso al modelo de objetos de proyecto VBA
# ═══════════════════════════════════════════════════════════════════════════

$ErrorActionPreference = "Stop"
$PSScriptRoot_safe     = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   Bot Boletas SII — Crear Template VBA    " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# ── Verificar Excel ──────────────────────────────────────────────────────
Write-Host "Iniciando Microsoft Excel..." -ForegroundColor Yellow
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
} catch {
    Write-Host ""
    Write-Host "✗ Excel no encontrado o no accesible." -ForegroundColor Red
    Write-Host "  Asegúrate de tener Microsoft Excel instalado." -ForegroundColor Red
    pause
    exit 1
}

$excel.Visible       = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # ── Módulo estándar con macros de utilidad ───────────────────────────
    $stdModule      = $wb.VBProject.VBComponents.Add(1)   # 1 = vbext_ct_StdModule
    $stdModule.Name = "MacrosBoletas"

    $vbaCode = @"
' ═══════════════════════════════════════════════════════════════
'  Bot Boletas SII — Macros de Utilidad
'  Accesibles via Alt+F8 o la cinta de opciones de macros.
' ═══════════════════════════════════════════════════════════════

' Ajusta el ancho de todas las columnas en todas las hojas
Sub AutoAjustarColumnas()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
    Next ws
    Application.ScreenUpdating = True
    MsgBox "Columnas ajustadas en " & ThisWorkbook.Worksheets.Count & " hojas.", _
           vbInformation, "Bot Boletas SII"
End Sub

' Navega directamente a la hoja Resumen
Sub IrAResumen()
    On Error Resume Next
    ThisWorkbook.Sheets("Resumen").Activate
    If Err.Number <> 0 Then
        MsgBox "No se encontró la hoja 'Resumen'.", vbExclamation, "Bot Boletas SII"
    End If
    On Error GoTo 0
End Sub

' Resalta en amarillo las celdas numéricas mayores al umbral ingresado
Sub ResaltarMontos()
    Dim sInput  As String
    Dim umbral  As Long
    Dim ws      As Worksheet
    Dim cel     As Range
    Dim cuenta  As Long

    sInput = InputBox("Ingrese el monto umbral (en $):" & Chr(10) & _
                      "Las celdas mayores a ese valor serán resaltadas.", _
                      "Resaltar Montos", "1000000")
    If sInput = "" Then Exit Sub
    umbral = Val(Replace(sInput, ".", ""))
    If umbral = 0 Then Exit Sub

    Application.ScreenUpdating = False
    cuenta = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Resumen" Then
            For Each cel In ws.UsedRange
                If IsNumeric(cel.Value) And cel.Value > umbral Then
                    cel.Interior.Color = RGB(255, 213, 79)
                    cel.Font.Bold      = True
                    cel.Font.Color     = RGB(90, 45, 0)
                    cuenta = cuenta + 1
                End If
            Next cel
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox cuenta & " celda(s) resaltada(s) con monto > $" & Format(umbral, "#,##0") & ".", _
           vbInformation, "Bot Boletas SII"
End Sub

' Quita todo el resaltado aplicado por ResaltarMontos
Sub LimpiarResaltado()
    Dim resp As Integer
    resp = MsgBox("¿Quitar todo el resaltado en el libro?", _
                  vbYesNo + vbQuestion, "Bot Boletas SII")
    If resp = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws.UsedRange
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex      = xlAutomatic
            .Font.Bold            = False
        End With
    Next ws
    Application.ScreenUpdating = True
End Sub

' Imprime todas las hojas de empresa (excluye Resumen)
Sub ImprimirEmpresas()
    Dim resp As Integer
    resp = MsgBox("Se imprimirán todas las hojas de empresa (excepto Resumen)." & Chr(10) & _
                  "¿Continuar?", vbYesNo + vbQuestion, "Bot Boletas SII")
    If resp = vbNo Then Exit Sub

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Resumen" Then
            ws.PrintOut Copies:=1
        End If
    Next ws
End Sub

' Muestra un resumen rápido de los totales del libro
Sub MostrarResumen()
    Dim ws     As Worksheet
    Dim msg    As String
    Dim found  As Boolean
    found = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Resumen" Then
            found = True
            Exit For
        End If
    Next ws

    If Not found Then
        MsgBox "No se encontró la hoja 'Resumen'.", vbExclamation, "Bot Boletas SII"
        Exit Sub
    End If

    msg = "Hojas de empresa: " & (ThisWorkbook.Worksheets.Count - 1) & Chr(10)
    msg = msg & "Haz clic en OK para ir al Resumen."
    MsgBox msg, vbInformation, "Bot Boletas SII"
    ThisWorkbook.Sheets("Resumen").Activate
End Sub
"@

    $stdModule.CodeModule.AddFromString($vbaCode)

    # ── Evento Workbook_Open en ThisWorkbook ─────────────────────────────
    $thisWbComp = $wb.VBProject.VBComponents("ThisWorkbook")
    $openCode   = @"
' Se ejecuta automáticamente al abrir el archivo
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    ' 1. Auto-ajustar columnas
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
    Next ws

    Application.ScreenUpdating = True

    ' 2. Ir a hoja Resumen si existe
    On Error Resume Next
    ThisWorkbook.Sheets("Resumen").Activate
    On Error GoTo 0
End Sub
"@

    $thisWbComp.CodeModule.AddFromString($openCode)

    # ── Guardar como xlsm ────────────────────────────────────────────────
    $templatesDir = Join-Path $PSScriptRoot_safe "..\templates"
    if (-not (Test-Path $templatesDir)) {
        New-Item -ItemType Directory -Path $templatesDir | Out-Null
    }

    $outPath = (Resolve-Path $templatesDir).Path + "\vba_template.xlsm"

    # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.SaveAs($outPath, 52)
    $wb.Close($false)

    Write-Host "✓ Template creado correctamente:" -ForegroundColor Green
    Write-Host "  $outPath" -ForegroundColor White
    Write-Host ""
    Write-Host "Macros incluidas (Alt+F8 para ejecutar):" -ForegroundColor Cyan
    Write-Host "  • AutoAjustarColumnas   Ajusta el ancho de todas las columnas"
    Write-Host "  • IrAResumen            Navega a la hoja Resumen"
    Write-Host "  • ResaltarMontos        Resalta celdas mayores a un umbral"
    Write-Host "  • LimpiarResaltado      Quita el resaltado aplicado"
    Write-Host "  • ImprimirEmpresas      Imprime todas las hojas de empresa"
    Write-Host "  • MostrarResumen        Muestra info rápida del libro"
    Write-Host ""
    Write-Host "Al abrir el .xlsm exportado se ejecutará automáticamente:" -ForegroundColor Yellow
    Write-Host "  • Auto-ajuste de columnas"
    Write-Host "  • Navegación a hoja Resumen"
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Si el error menciona VBProject o seguridad, haz esto en Excel:" -ForegroundColor Yellow
    Write-Host "  Archivo → Opciones → Centro de confianza"
    Write-Host "  → Configuración del centro de confianza → Configuración de macros"
    Write-Host "  → Marcar: 'Confiar en el acceso al modelo de objetos de proyecto VBA'"
    Write-Host ""
} finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}

Write-Host ""
pause
