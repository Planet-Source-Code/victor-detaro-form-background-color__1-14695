Attribute VB_Name = "Module1"
'* BACKGROUND COLOR CHANGE made by Victor Detaro
'* Module to change background color of form from
'* one color to the other
'*
'* FEEL FREE TO USE OR MODIFY,
'* BUT DON'T FORGET TO INCLUDE MY NAME IN YOUR CREDITS IF YOU DO USE

Function bgcolor(frm As Form, rt As Integer, gt As Integer, bt As Integer, rb As Integer, gb As Integer, bb As Integer)
    'INITIALIZATION OF FORM GRAPHIC DETAILS
    If frm.WindowState = 1 Then GoTo vic
    frm.DrawStyle = 6
    frm.ScaleMode = 3
    frm.AutoRedraw = True
    frm.DrawMode = 13
    
    'SCALING FORM TO FIT THE STANDARD COLOR RANGE
    If frm.ScaleWidth <> 255 Then
        frm.ScaleWidth = 255
    End If
    If frm.ScaleHeight <> 255 Then
        frm.ScaleHeight = 255
    End If
    
    'CALCULATING THE No. OF LINES TO FILL
    frm.DrawWidth = frm.ScaleHeight / 255
    j = 0
    frm.Refresh
    
    'INITIALIZING TOP COLOR
    m = rt
    n = gt
    o = bt
    
    'TOP COLOR -> BOTTOM COLOR
    For i = 0 To 255
        frm.Line (0, i)-(frm.Width, i + frm.DrawWidth), RGB(m, n, o), BF
        j = j + 1
        If rt < rb And m < rb Then
            m = m + 1
        ElseIf rt > rb And m > rb Then
            m = m - 1
        End If
        If gt < gb And n < gb Then
            n = n + 1
        ElseIf gt > gb And n > gb Then
            n = n - 1
        End If
        If bt < bb And o < bb Then
            o = o + 1
        ElseIf bt > bb And o > bb Then
            o = o - 1
        End If
    Next i
vic:
End Function
