Attribute VB_Name = "Set_PPT_Font"
Dim pptSlide As Slide
Dim pptShape As Shape
Dim pptTable As Table

Dim CNFont As String
Dim EnFont As String
Sub main()
    'Move your cursor to here, and click F5, it will change your font in all pages incluing shapes, tables, but font in smart art shape are excluded.
    Call SetShapeFont
End Sub
Sub SetShapeFont()
    Dim i As Long
    Dim J As Long
    
    CNFont = "Microsoft YaHei" 'Font for Chinese
    EnFont = "Microsoft YaHei" 'Font for English
    
    With ActivePresentation
        For Each pptSlide In .Slides
            For Each pptShape In pptSlide.Shapes
                With pptShape
                    If .HasTextFrame Then
                        If .TextFrame.HasText Then
                            'font for Chinese and Japanse
                            .TextFrame.TextRange.Font.NameFarEast = CNFont
                            'font for english
                            .TextFrame.TextRange.Font.Name = EnFont
                        End If
                    End If
                End With
                If pptShape.HasTable Then
                    Set pptTable = pptShape.Table
                    For i = 1 To pptTable.Columns.Count
                        For J = 1 To pptTable.Rows.Count
                            With pptTable.Cell(J, i).Shape.TextFrame.TextRange.Font
                                .Name = EnFont
                                .NameFarEast = CNFont
                            End With
                        Next J
                    Next i
                End If
            Next
        Next
    End With
End Sub

