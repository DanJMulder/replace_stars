'   ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ##
'   ##   Macro for turning stars in word document into commonly repeated text     ##
'   ##   Written by Daniel Mulder, August 2022.                                   ##
'   ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ## ##
  
' This is a visual basic script that makes a macro to apply to my clinical note templates to remove/replace the "*" patterns in a word doc with very common words, to save time cursoring to these spots and changing the text manually
  
' Key: 
' ***** = normal
' **** = none
' *** = no
' ** = [blank]


Sub ReplaceStars()

With Selection.Find
 .ClearFormatting
 .Text = "*****"
 .Replacement.ClearFormatting
 .Replacement.Text = "normal"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With

With Selection.Find
 .ClearFormatting
 .Text = "****"
 .Replacement.ClearFormatting
 .Replacement.Text = "none"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With

With Selection.Find
 .ClearFormatting
 .Text = "***"
 .Replacement.ClearFormatting
 .Replacement.Text = "no"
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With

With Selection.Find
 .ClearFormatting
 .Text = "** "
 .Replacement.ClearFormatting
 .Replacement.Text = ""
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With

With Selection.Find
 .ClearFormatting
 .Text = "**"
 .Replacement.ClearFormatting
 .Replacement.Text = ""
 .Execute Replace:=wdReplaceAll, Forward:=True, _
 Wrap:=wdFindContinue
End With

End Sub
