Attribute VB_Name = "modCaption"
Option Explicit

'*********************************************'
'*        Caption Alignment Constants        *'
'*********************************************'
Private Const TOPLEFT As Long = 0
Private Const TOPCENTER As Long = 1
Private Const TOPRIGHT As Long = 2
Private Const BOTTOMLEFT As Long = 3
Private Const BOTTOMCENTER As Long = 4
Private Const BOTTOMRIGHT As Long = 5

'*********************************************'
'*             3D Font Constants             *'
'*********************************************'
Private Const RAISEDLIGHT As Long = 1
Private Const RAISEDHEAVY As Long = 2
Private Const INSETLIGHT As Long = 3
Private Const INSETHEAVY As Long = 4
Private Const DROPSHADOW As Long = 5

'*********************************************'
'*         DrawText uFlags Constants         *'
'*********************************************'
Private Const DT_TOP = &H0               ' Top-justifies text (single line only).
Private Const DT_LEFT = &H0              ' Aligns text to the left.
Private Const DT_CENTER = &H1            ' Centers text horizontally in the rectangle.
Private Const DT_RIGHT = &H2             ' Aligns text to the right.
Private Const DT_VCENTER = &H4           ' Centers text vertically (single line only).
Private Const DT_BOTTOM = &H8            ' Justifies text to the bottom of the rectangle. This value must be combined with DT_SINGLELINE.
Private Const DT_WORDBREAK = &H10        ' Breaks words. Lines are automatically broken between words if a word would extend past the edge of the rectangle specified by the DestRect parameter. A carriage return/line feed sequence also breaks the line.
Private Const DT_SINGLELINE = &H20       ' Displays text on a single line only. Carriage returns and line feeds do not break the line.
Private Const DT_EXPANDTABS = &H40       ' Expands tab characters. The default number of characters per tab is eight. The DT_WORD_ELLIPSIS, DT_PATH_ELLIPSIS, and DT_END_ELLIPSIS values cannot be used with the DT_EXPANDTABS value.
Private Const DT_TABSTOP = &H80          ' Sets tab stops. Bits 15–8 (high-order byte of the low-order word) of the Format parameter specify the number of characters for each tab. The default number of characters per tab is eight. The DT_CALCRECT, DT_EXTERNALLEADING, DT_INTERNAL, DT_NOCLIP, and DT_NOPREFIX values cannot be used with the DT_TABSTOP value.
Private Const DT_NOCLIP = &H100          ' Draws without clipping. DrawTextW is somewhat faster when DT_NOCLIP is used.
Private Const DT_EXTERNALLEADING = &H200 ' Includes the font external leading in line height. Normally, external leading is not included in the height of a line of text.
Private Const DT_CALCRECT = &H400        ' Determines the width and height of the rectangle. If there are multiple lines of text, DrawTextW uses the width of the rectangle pointed to by the SrcRect parameter and extends the base of the rectangle to bound the last line of text. If there is only one line of text, DrawTextW modifies the right side of the rectangle so that it bounds the last character in the line. In either case, DrawTextW returns the height of the formatted text but does not draw the text.
Private Const DT_NOPREFIX = &H800        ' Turns off processing of prefix characters. Normally, DrawTextW interprets the mnemonic-prefix character ampersand (&) as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off. Compare with DT_HIDEPREFIX and DT_PREFIXONLY.
Private Const DT_INTERNAL = &H1000       ' Uses the system font to calculate text metrics.
Private Const DT_EDITCONTROL = &H2000    ' Duplicates the text-displaying characteristics of a multiline edit control. Specifically, the average character width is calculated in the same manner as for an edit control, and the function does not display a partially visible last line.
Private Const DT_PATH_ELLIPSIS = &H4000  ' For displayed text, replaces characters in the middle of the string with ellipses so that the result fits in the specified rectangle. If the string contains backslash (\) characters, DT_PATH_ELLIPSIS preserves as much as possible of the text after the last backslash.
Private Const DT_END_ELLIPSIS = &H8000   ' For displayed text, if the end of a string does not fit in the rectangle, it is truncated and ellipses are added. If a word that is not at the end of the string goes beyond the limits of the rectangle, it is truncated without ellipses. The string is not modified unless the DT_MODIFYSTRING flag is specified. Compare with DT_PATH_ELLIPSIS and DT_WORD_ELLIPSIS.
Private Const DT_MODIFYSTRING = &H10000  ' Modifies the specified string to match the displayed text. This value has no effect unless DT_END_ELLIPSIS or DT_PATH_ELLIPSIS is specified.
Private Const DT_RTLREADING = &H20000    ' Layout in right-to-left reading order for bi-directional text when the font selected into the hdc is a Hebrew or Arabic font. The default reading order for all text is left-to-right.
Private Const DT_WORD_ELLIPSIS = &H40000 ' Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
Private Const DT_HIDEPREFIX = &H100000   ' Windows 2000/XP: Ignores the ampersand (&) prefix character in the text. The letter that follows will not be underlined, but other mnemonic-prefix characters are still processed. Compare with DT_NOPREFIX and DT_PREFIXONLY.
Private Const DT_PREFIXONLY = &H200000   ' Turns off processing of prefix characters. Normally, DrawText interprets the mnemonic-prefix character & as a directive to underscore the character that follows, and the mnemonic-prefix characters && as a directive to print a single &. By specifying DT_NOPREFIX, this processing is turned off. Compare with DT_HIDEPREFIX and DT_PREFIXONLY.
Private Const DT_NOFULLWIDTHCHARBREAK = &H80000  ' Windows 98/Me, Windows 2000/XP: Prevents a line break at a DBCS (double-wide character string), so that the line breaking rule is equivalent to SBCS strings. For example, this can be used in Korean windows, for more readability of icon labels. This value has no effect unless DT_WORDBREAK is specified.

'***********************************************'
'*              Win32 API Declares             *'
'***********************************************'
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Function GetCaptionHeight(ByVal Target As Control, _
                                 ByVal sCaption As String, _
                                 Optional lWordWrap As Long = 0) As RECT

Dim rcTemp As RECT
Dim DSFlag As Long

' Set the left and right boundaries of the client area of the
' caption.  We calculate the width and height of the text based
' on this area, and gives us a start and end point for calculating
' word wrapped text.
GetClientRect Target.hWnd, rcTemp
rcTemp.Left = rcTemp.Left + 7
rcTemp.Right = rcTemp.Right - 7

' Decide if we are using word wrapped captions or not.
DSFlag = IIf(lWordWrap, DT_WORDBREAK, DT_SINGLELINE) Or DT_CALCRECT

' Calculate the actual widhth and height of the caption.
' nCount of -1 assumes that sCaption is a Null string and DrawText
' computes the character count automatically.
DrawText Target.hDC, sCaption, Len(sCaption), rcTemp, DSFlag

GetCaptionHeight = rcTemp

End Function

Public Sub DrawCaption(ByVal hDestDC As Long, ByRef Target As Control, _
                       ByVal sCaption As String, ByRef rcClient As RECT, _
                       Optional bEnabled As Boolean = True, _
                       Optional lFont3D As Long = 0, _
                       Optional lAlign As Long = 0)

Dim DSFlag As Long
Dim lColorOld As OLE_COLOR

lColorOld = Target.ForeColor

DSFlag = DT_WORDBREAK

Select Case lAlign
    Case 0, 1, 2
        DSFlag = DSFlag Or DT_LEFT
    Case 3, 4, 5
        DSFlag = DSFlag Or DT_RIGHT
End Select

If bEnabled = True Then
    Select Case lFont3D
        Case RAISEDLIGHT
            OffsetRect rcClient, -1, -1
            Target.ForeColor = vb3DHighlight
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, 1, 1
                    
        Case RAISEDHEAVY
            OffsetRect rcClient, -1, -1
            Target.ForeColor = vb3DHighlight
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, 2, 2
            Target.ForeColor = vb3DShadow
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, -1, -1
        
        Case INSETLIGHT
            OffsetRect rcClient, 1, 1
            Target.ForeColor = vb3DHighlight
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, -1, -1
                    
        Case INSETHEAVY
            OffsetRect rcClient, -1, -1
            Target.ForeColor = vb3DShadow
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, 2, 2
            Target.ForeColor = vb3DHighlight
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, -1, -1
        
        Case DROPSHADOW
            OffsetRect rcClient, 1, 1
            Target.ForeColor = vb3DShadow
            DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
            OffsetRect rcClient, -1, -1
            
    End Select
    Target.ForeColor = lColorOld
    DrawText Target.hDC, sCaption, -1, rcClient, DSFlag
Else
    ' Store the old color.
    'lColorOld = GetTextColor(Target.hDC)
    
    SetTextColor Target.hDC, TranslateColor(vb3DHighlight)
    OffsetRect rcClient, 1, 1
    DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
    
    SetTextColor Target.hDC, TranslateColor(vb3DShadow)
    OffsetRect rcClient, -1, -1
    DrawText Target.hDC, sCaption, Len(sCaption), rcClient, DSFlag
    
    ' Restore the old color.
    SetTextColor Target.hDC, TranslateColor(lColorOld)
End If

BitBlt hDestDC, rcClient.Left, rcClient.Top, rcClient.Right - rcClient.Left, rcClient.Bottom, Target.hDC, rcClient.Left, rcClient.Top, vbSrcCopy

End Sub

