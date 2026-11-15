Attribute VB_Name = "CertusVariantes"
' Certus PLD - Resolver variantes PF / PM en documentos del kit
' Estilos usados: "Variante PF", "Variante PM", "Nota Certus"
Option Explicit

Public Sub ResolverComoPF()
    Resolver "Variante PF", "Variante PM", "[PF: "
End Sub

Public Sub ResolverComoPM()
    Resolver "Variante PM", "Variante PF", "[PM: "
End Sub

Private Sub Resolver(conservar As String, eliminar As String, etiqueta As String)
    If MsgBox("Se eliminaran las variantes que no aplican y la nota de edicion. " & _
              "Conviene guardar antes una copia del documento. ¿Continuar?", _
              vbYesNo + vbQuestion, "Certus PLD") = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    BorrarPorEstilo "Nota Certus"
    BorrarPorEstilo eliminar
    ReemplazarEnEstilo conservar, etiqueta, ""
    ReemplazarEnEstilo conservar, "]", ""
    QuitarEstilo conservar
    Application.ScreenUpdating = True
    MsgBox "Documento resuelto. Revise la redaccion antes de entregarlo.", vbInformation, "Certus PLD"
End Sub

Private Sub BorrarPorEstilo(nombre As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "": .Replacement.Text = ""
        .Style = ActiveDocument.Styles(nombre)
        .Format = True: .MatchWildcards = False: .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub ReemplazarEnEstilo(nombre As String, buscar As String, reemplazo As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = buscar: .Replacement.Text = reemplazo
        .Style = ActiveDocument.Styles(nombre)
        .Format = True: .MatchWildcards = False: .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub QuitarEstilo(nombre As String)
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "": .Replacement.Text = ""
        .Style = ActiveDocument.Styles(nombre)
        .Replacement.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
        .Format = True: .MatchWildcards = False: .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
