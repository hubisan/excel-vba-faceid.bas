Attribute VB_Name = "faceid"
'This content is released under the (http://opensource.org/licenses/MIT) MIT License.
'Copyright (c) Daniel Hubmann (hubisan@gmail.com)

'*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸

'Module: faceid
'Add faceids to worksheet with number below
'Not turning screenupdating off as this gave me some weird errors.

'*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸¸.•*´¨*•.¸

Option Explicit

'example for calling the main sub
'the output of this is already generated and visible in the worksheet "faceid"
Sub call_faceid()
    Call faceids_to_wks(ThisWorkbook.Worksheets("faceid"), "B2", 1, 1000, 25)
End Sub

'add faceids with image and number to worksheet
'clears the worksheet and sets heights, withs and fonts
'@param wks wks: worksheet to print faceids on
'@param first_cell rng: range of cell with first faceid
'@param min_face_id lng: first faceid to print
'@param max_face_id lng: last faceid to print (will print from min to max)
'@param ids_per_row lng: ids per row
Sub faceids_to_wks(wks As Worksheet, str_first_cell As String, _
min_face_id As Long, max_face_id As Long, ids_per_row As Long)
    Dim popup_menu As CommandBar
    Dim face_id As Long
    Dim first_cell As Range
    Dim pic As Object
    Dim rng As Range
    
    On Error GoTo error_handler
    
    If min_face_id > max_face_id Then Exit Sub
    Set first_cell = wks.Range(str_first_cell)
    face_id = min_face_id
    
    'tabula rasa
    With wks.Cells
        .Clear
        .Font.Size = 6
        .Font.Name = "Calibri Light"
        .VerticalAlignment = xlBottom
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
        .ColumnWidth = 3
    End With
    wks.DrawingObjects.Delete
    'wks.Shapes.SelectAll: Selection.Delete 'OLD
    
    'loop through a range with max_ids cells respecting the ids_per_row variable
    For Each rng In wks.Range(first_cell, wks.Cells(wks.Rows.Count, first_cell.Offset(0, ids_per_row - 1).Column))
        Set popup_menu = Application.CommandBars.Add("temp_face", msoBarPopup, False, True)
        With popup_menu.Controls.Add(Type:=msoControlButton)
            .faceid = face_id
            .CopyFace
        End With
        wks.Paste Destination:=rng
        Set pic = Selection
        With pic
            .Name = "faceid_" & face_id
            .Left = .Left + 4.5
            .Top = .Top + 4.5
        End With
        rng.Value = face_id
        rng.Select 'to deselect the picture
        Application.CommandBars("temp_face").Delete
        Set popup_menu = Nothing
        face_id = face_id + 1
        If face_id > max_face_id Then Exit For
    Next
    
exit_sub:
    On Error Resume Next
    Application.CommandBars("temp_face").Delete
    Exit Sub
error_handler:
    MsgBox Err.Description
    Resume exit_sub
End Sub
