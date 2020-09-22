VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Light RTF editor class
'Written by adi barda israel (adib@malam.com)

'This is a light version of my editor class
'The complete class also supports VB like intelisence


Option Explicit


Private Type InterfaceData
    Class As String
    method As String
    IsMethod As Boolean
    ToolTip As String
End Type
Private m_Interface() As InterfaceData

Const abGREEN = 32768

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type EditorWords
    color As Long
    wrd As String
End Type
Private m_arrWords() As EditorWords

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private m_bPaintText As Boolean

Private m_iPicHeight As Long
Private WithEvents m_txtScript As RichTextLib.RichTextBox
Attribute m_txtScript.VB_VarHelpID = -1
Private m_DebugTextbox As RichTextLib.RichTextBox

Public Sub AddEditorWord(ByVal wrd As String, ByVal color As Long)

    Dim i  As Long
    
    i = UBound(m_arrWords())
    
    If m_arrWords(0).wrd = "" Then
        i = 0
    Else
        i = i + 1
    End If
    
    ReDim Preserve m_arrWords(i)
    m_arrWords(i).wrd = wrd
    m_arrWords(i).color = color
    
    
End Sub

Public Sub SetEditorObjects(ByRef objRTF As RichTextLib.RichTextBox, _
                            ByRef DebugTextbox As RichTextLib.RichTextBox)
                           
    'Init editor objects
    Set m_txtScript = objRTF 'main script text box
    Set m_DebugTextbox = DebugTextbox 'need this for smooth painting
    
    ClearScript
    
End Sub

Public Sub PaintText()

    Dim iPos As Long
    Dim i As Long
    Dim iCnt As Long
    
    Dim arrWrd() As String
    
    
   
    ReDim arrWrd(0)
    iCnt = 0
    
    iPos = m_txtScript.SelStart
    
    For i = 0 To UBound(m_arrWords())
    
        If m_arrWords(i).wrd = "" Then
            Exit For
        End If
        
        ReDim Preserve arrWrd(iCnt)
        arrWrd(iCnt) = m_arrWords(i).wrd
        
        'Is it the last word to paint ?
        If UBound(m_arrWords()) = i Then
            'Yes it is
            ColorWord m_txtScript, arrWrd(), m_arrWords(i).color
        Else
            If m_arrWords(i + 1).color <> m_arrWords(i).color Then
                iCnt = 0
                ColorWord m_txtScript, arrWrd(), m_arrWords(i).color
            Else
                iCnt = iCnt + 1
            End If
        End If
        
    Next i
    
    
    ColorStrings m_txtScript, Chr$(34), vbMagenta
    ColorRow m_txtScript, "'", abGREEN
    
    
    
    If iPos > 0 Then
        iPos = InStr(iPos, m_txtScript.Text, Chr$(10))
    End If
    m_txtScript.SelStart = iPos
    
End Sub

Public Sub ColorWord(ByVal objRTF As RichTextLib.RichTextBox, ByRef wrd() As String, ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim sRTF As String
    Dim iColor As Long
    
    With objRTF
    
        sRTF = .TextRTF
        
        For i = 0 To UBound(wrd)
            
            iStart = InStr(1, sRTF, wrd(i), vbTextCompare)
            If iStart > 0 Then
                'Check && update the color table
                iColor = SetColorTable(sRTF, color)
                
                sRTF = Replace$(sRTF, wrd(i), "\cf" & iColor & " " & wrd(i) & "\cf0", , , vbTextCompare)
                
            End If
                
            
        Next i
        
        .TextRTF = sRTF
        
    End With
    
End Sub

Private Function SetColorTable(ByRef rtf As String, ByVal color As Long) As Long

    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sTable As String
    Dim sTmp As String
    Dim sColor As String
    
    GetRGB color, iR, iG, iB
    sColor = "\red" & iR & "\green" & iG & "\blue" & iB & ";"
    
    'Do we have color table?
    iStart = InStr(1, rtf, "colortbl", vbTextCompare)
    If iStart = 0 Then
        'No, Insert color table
        rtf = Replace$(rtf, "}}", "}}{\colortbl ;}", , 1)
        iStart = InStr(1, rtf, "colortbl", vbTextCompare)
    End If
    
    'Do we have the current color?
    If InStr(1, rtf, sColor, vbTextCompare) = 0 Then
        'No, insert color
        i = InStr(1, rtf, "\colortbl")
        i = InStr(i, rtf, "}", vbTextCompare)
        sTmp = Left$(rtf, i - 1)
        rtf = Replace$(rtf, "}", sColor & "}", i, 1, vbTextCompare)
        rtf = sTmp & rtf
    End If
    
    iEnd = InStr(iStart, rtf, ";}")
    sTable = Mid$(rtf, iStart, iEnd - iStart + 1)
    
    i = 0
    iStart = 1
    Do
        iStart = InStr(iStart, sTable, ";")
        If iStart <> 0 Then
            i = i + 1
            iStart = iStart + 1
            If InStr(iStart, sTable, sColor, vbTextCompare) = iStart Then
                SetColorTable = i
                Exit Do
            End If
        End If
    Loop Until iStart = 0
    
End Function

Private Sub GetRGB(ByVal color As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    
    b = Int(color / 65536)
    g = Int((color / 65536 - b) * 65536 / 256)
    r = Int(((color / 65536 - b) * 65536 / 256 - g) * 256)
    
End Sub

Private Sub ColorRow(ByVal objRTF As RichTextLib.RichTextBox, ByVal wrd As String, ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sRTF As String
    Dim sRow As String
    Dim sNewRow As String
    Dim iColor As Long
    Dim iCommaPlace As Long
    
    With objRTF
        
        sRTF = .TextRTF
        
        iStart = 1
        Do
            iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
            If iStart > 0 Then
            
                'Check if its hebrew letter
                If Mid$(sRTF, iStart - 1, 1) <> "\" Then
                
                    'Check && update the color table
                    iColor = SetColorTable(sRTF, color)
                    
                    iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
                    iEnd = InStr(iStart, sRTF, vbNewLine)
                    iCommaPlace = InStr(iStart, sRTF, Chr$(34))
                    
                    If iEnd > 0 And (iCommaPlace = 0 Or iCommaPlace > iEnd) Then
                        sRow = Mid$(sRTF, iStart, iEnd - iStart)
                        sNewRow = GetCleanRow(sRow)
                        sRTF = Replace$(sRTF, sRow, "\cf" & iColor & " " & sNewRow & "\cf0", , , vbTextCompare)
                        iEnd = InStr(iStart, sRTF, vbNewLine)
                    End If
                    
                    
                    iStart = iEnd
                    
                Else
                
                    iStart = iStart + 1
                    
                End If 'Not hebrew letter
                
            End If 'iStart>0
            
        Loop Until iStart = 0
        
        .TextRTF = sRTF
        
    End With
    

End Sub

Private Sub ColorStrings(ByVal objRTF As RichTextLib.RichTextBox, ByVal wrd As String, ByVal color As Long)

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim iBreak As Long
    Dim sRTF As String
    Dim sRow As String
    Dim sNewRow As String
    Dim iColor As Long
    
    With objRTF
        
        sRTF = .TextRTF
        
        iStart = 1
        Do
            iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
            If iStart > 0 Then
            
                'Check && update the color table
                iColor = SetColorTable(sRTF, color)
                
                iStart = InStr(iStart, sRTF, wrd, vbTextCompare)
                iEnd = InStr(iStart + 1, sRTF, wrd, vbTextCompare)
                iBreak = InStr(iStart + 1, sRTF, vbNewLine)
                
                If (iEnd > 0) And ((iBreak > iEnd) Or (iBreak = 0)) Then
                    
                    sRow = Mid$(sRTF, iStart, iEnd - iStart + 1)
                    'sNewRow = GetCleanRow(sRow)
                    sNewRow = sRow 'just for backward compatibility
                    sRTF = Replace$(sRTF, sRow, "\cf" & iColor & " " & sNewRow & "\cf0", , , vbTextCompare)
                    'iEnd = InStr(iStart, sRTF, vbNewLine)
                    ' \ltrch \'20
                    
                    iEnd = InStr(iEnd + Len("\cf" & iColor & " " & sNewRow & "\cf0") - Len(sNewRow), sRTF, wrd)
                End If
                
                
                iStart = iEnd
                
                
            End If 'iStart>0
            
        Loop Until iStart = 0
        
        .TextRTF = sRTF
        
    End With
    

End Sub


Private Function GetCleanRow(ByVal Row As String) As String

    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim sTmp As String
    
    GetCleanRow = Row
    iStart = 1
    Do
        iStart = InStr(iStart, Row, "\")
        If iStart > 0 Then
            iEnd = InStr(iStart, Row, " ")
            If iStart > 0 Then
                If iEnd = 0 Then
                    iEnd = Len(Row) + 1
                Else
                    iEnd = iEnd + 1
                End If
                sTmp = Mid$(Row, iStart, iEnd - iStart)
                Row = Replace$(Row, sTmp, "")
            End If
        End If
        
    Loop Until iStart = 0
    
    GetCleanRow = Row
    
End Function

Public Sub ClearScript()

    m_txtScript.Text = ""

End Sub


Private Sub Class_Initialize()
    
    m_bPaintText = False
    
    ReDim m_arrWords(0)
    
End Sub



Private Sub m_txtScript_Change()

    If m_bPaintText Then
        m_bPaintText = False
        
        m_DebugTextbox.SetFocus
        PaintText
        m_txtScript.SetFocus

    End If
    
End Sub

Private Sub m_txtScript_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn ', vbKeySpace
            m_bPaintText = True
        Case vbKeyTab
            KeyCode = 0
            m_txtScript.SelText = "    "
            DoEvents
            m_txtScript.SetFocus
    End Select

End Sub
