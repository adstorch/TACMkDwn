For Each s In ActiveDocument.Sections
  With s.Headers(wdHeaderFooterPrimary).PageNumbers
    .RestartNumberingAtSection = False
    .StartingNumber = 0
	End With
Next s
With ActiveDocument
    .PageSetup.DifferentFirstPageHeaderFooter = True
End With
    Dim oStory As Range
    For Each oStory In ActiveDocument.StoryRanges
        oStory.Fields.Update
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
                Set oStory = oStory.NextStoryRange
                oStory.Fields.Update
            Wend
        End If
    Next oStory
    Set oStory = Nothing

    Dim TOC As TableOfContents
    For Each TOC In ActiveDocument.TablesOfContents
        TOC.Update
    Next