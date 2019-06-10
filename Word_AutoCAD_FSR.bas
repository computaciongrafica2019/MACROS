Attribute VB_Name = "Módulo1"
Sub Word_AutoCAD()
'Fernando Sandoval
Dim acadApp As AcadApplication
Set acadApp = New AcadApplication
'Set acadApp = AcadApplication
acadApp.Documents.Add 'Añade nuevo dibujo en autocad
'acadApp.Documents.Open "C:\Users\Fernando\Documents\Fer\Semestre 2019-1\Computacion grafica\SALACAD.dwg"
'abre dibujo de AutoCAD, debe buscar la ruta
'acadApp.Documents.Open "C:\Users\Fernando\Documents\Fer\Semestre 2019-1\Computacion grafica\CODIGOS\Drawing1.dwg"
acadApp.ActiveDocument.SendCommand "_Circle" & vbCr & "0,0,0" & vbCr & "4" & vbCr
'acadApp.ActiveDocument.SendCommand "Vlisp" & vbCr
Dim Centro(1 To 3) As Double
Centro(1) = 100: Centro(2) = 100: Centro(3) = 0
Dim Radio As Double
Radio = 50
Set circulo = acadApp.ActiveDocument.ModelSpace.AddCircle(Centro, Radio)
'acadApp.ActiveDocument.SendCommand "(load ""C:\\Users\\Fernando\\Documents\\Fer\\Semestre 2019-1\\Computacion grafica\\CODIGOS\\Reloj3V_digital_section.lsp"")" & vbCr
End Sub

