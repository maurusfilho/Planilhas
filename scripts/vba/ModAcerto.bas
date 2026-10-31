Option Explicit

' ============================================================
'  HLM - Acerto de contas mensal
'  Botao "BAIXAR ACERTO (PDF)" da aba ACERTO MENSAL.
'  Gera um PDF somente com a aba ACERTO MENSAL do mes
'  selecionado, escondendo as linhas vazias do detalhamento.
' ============================================================

Public Sub BaixarAcertoPDF()
    Dim ws As Worksheet
    Dim mesRef As Variant
    Dim nomeArq As String, emp As String
    Dim destino As Variant
    Dim ocultas As Collection

    On Error GoTo Falha
    Set ws = ThisWorkbook.Worksheets("ACERTO MENSAL")
    mesRef = ws.Range("SelMes").Value
    If Not IsDate(mesRef) Then
        MsgBox "Selecione o mes do acerto (celula amarela) antes de baixar.", vbExclamation, "Acerto de contas"
        Exit Sub
    End If

    Application.Calculate
    emp = LimparNome(CStr(ThisWorkbook.Worksheets("CONFIG").Range("NomeEmp").Value))
    nomeArq = "Acerto_" & Format(mesRef, "yyyy-mm") & "_" & emp & ".pdf"

    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        destino = Application.GetSaveAsFilename(InitialFileName:=nomeArq)
    Else
        destino = Application.GetSaveAsFilename(InitialFileName:=nomeArq, _
                    FileFilter:="Arquivo PDF (*.pdf),*.pdf", Title:="Salvar acerto de contas do mes")
    End If
    If VarType(destino) = vbBoolean Then Exit Sub
    If LCase(Right(CStr(destino), 4)) <> ".pdf" Then destino = CStr(destino) & ".pdf"

    Set ocultas = New Collection
    Application.ScreenUpdating = False
    OcultarVazias ws, "AcertoDetalhe", ocultas
    OcultarVazias ws, "AcertoFolha", ocultas

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(destino), _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True

    Reexibir ocultas
    Application.ScreenUpdating = True
    MsgBox "Acerto de contas salvo em:" & vbCrLf & CStr(destino), vbInformation, "Acerto de contas"
    Exit Sub

Falha:
    On Error Resume Next
    If Not ocultas Is Nothing Then Reexibir ocultas
    Application.ScreenUpdating = True
    MsgBox "Nao foi possivel gerar o PDF: " & Err.Description, vbCritical, "Acerto de contas"
End Sub

Private Sub OcultarVazias(ByVal ws As Worksheet, ByVal nomeFaixa As String, ByVal ocultas As Collection)
    Dim c As Range
    For Each c In ws.Range(nomeFaixa).Cells
        If CStr(c.Value) = "" Then
            If Not c.EntireRow.Hidden Then
                c.EntireRow.Hidden = True
                ocultas.Add c.Row
            End If
        End If
    Next c
End Sub

Private Sub Reexibir(ByVal ocultas As Collection)
    Dim ws As Worksheet
    Dim r As Variant
    Set ws = ThisWorkbook.Worksheets("ACERTO MENSAL")
    For Each r In ocultas
        ws.Rows(CLng(r)).Hidden = False
    Next r
End Sub

Private Function LimparNome(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "")
    Next c
    s = Replace(Trim(s), " ", "_")
    If Len(s) = 0 Then s = "Empreendimento"
    LimparNome = s
End Function
