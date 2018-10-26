Attribute VB_Name = "Referencer2000"
Sub DrawLineAtClick()

Dim doc As Document, retval As Long
Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, shift As Long
Dim numberValue As String


Set doc = ActiveDocument
doc.Unit = cdrCentimeter

retval = doc.GetUserClick(x1, y1, shift, 10, True, cdrCursorPick)
retval = doc.GetUserClick(x2, y2, shift, 10, True, cdrCursorPick)

InserirNumero.TextBox.Text = ""
InserirNumero.Show

If x2 - x1 >= 0 And y2 - y1 >= 0 Then
    Dim line1 As Shape
    Set line1 = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)

    If InserirNumero.TextBox.Text <> "" Then
    Dim TextBox1 As Shape
    Set TextBox1 = ActiveLayer.CreateArtisticText(x2, y2, InserirNumero.TextBox.Text, , , , 14)
    End If
    
Else
If x2 - x1 < 0 And y2 - y1 >= 0 Then
    Dim line2 As Shape
    Set line2 = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)

    If InserirNumero.TextBox.Text <> "" Then
    Dim TextBox2 As Shape
    Set TextBox2 = ActiveLayer.CreateArtisticText(x2 - 0.5, y2, InserirNumero.TextBox.Text, , , , 14)
    End If
        
Else
If x2 - x1 >= 0 And y2 - y1 < 0 Then
    Dim line3 As Shape
    Set line3 = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)

    If InserirNumero.TextBox.Text <> "" Then
    Dim TextBox3 As Shape
    Set TextBox3 = ActiveLayer.CreateArtisticText(x2, y2 - 0.4, InserirNumero.TextBox.Text, , , , 14)
    End If
    
Else
If x2 - x1 < 0 And y2 - y1 < 0 Then
    Dim line4 As Shape
    Set line4 = ActiveLayer.CreateLineSegment(x1, y1, x2, y2)
    
    If InserirNumero.TextBox.Text <> "" Then
    Dim TextBox4 As Shape
    Set TextBox4 = ActiveLayer.CreateArtisticText(x2 - 0.5, y2 - 0.4, InserirNumero.TextBox.Text, , , , 14)
    End If
            
End If
End If
End If
End If

End Sub
