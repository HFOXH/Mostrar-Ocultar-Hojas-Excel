# 📂 Mostrar-Ocultar-Hojas-Excel

Este repositorio contiene un libro y código VBA en macros sencillo para **mostrar y ocultar hojas en Excel utilizando hipervínculos**. En la primera hoja tienes los enlaces para ir a las otras hojas que estan ocultas y solo se habilitan si accedes desde ahi.  

## 🚀 Funcionalidad  
- **Ocultar** las hojas de excel y **mostrarlas** a través de un hipervinculo. 

## 📺 Tutorial
Mira este video explicativo en YouTube: <a href="https://www.youtube.com/watch?v=Lgq5x3g0h0I&lc=UgyFRyh5L_WDOfE-Tep4AaABAg">Tutorial</a>

```vba
' Este código permite ocultar y mostrar pestañas en Excel usando hipervínculos
' Debe colocarse en un módulo de VBA
' Tambien lo puedes encontrar en el Código.txt

Dim DirLink As String   'DIRECCION DEL HIPERVINCULO SELECCIONADO
 
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Sheets("PARAMETROS").Activate
       DirLink = ActiveCell.Hyperlinks(1).SubAddress
       DirLink = Left(DirLink, InStr(1, DirLink, "!", vbTextCompare) - 1)
    
       If Sheets(DirLink).Visible = False Then
           Sheets(DirLink).Visible = True
           Sheets(DirLink).Activate
           Sheets(DirLink).Range("A1").Select
       End If
End Sub
 
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    If Sh.Name <> "PARAMETROS" Then Sheets(Sh.Name).Visible = False
End Sub
