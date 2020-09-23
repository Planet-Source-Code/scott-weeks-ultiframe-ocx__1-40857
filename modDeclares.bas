Attribute VB_Name = "modDeclares"
Option Explicit

'*********************************************'
'*              Public Constants             *'
'*********************************************'
Public Const CLR_INVALID = &HFFFF

'*********************************************'
'*        PlaySound dwFlags Constants        *'
'*********************************************'
Public Const SND_SYNC = &H0             ' Synchronous playback of a sound event. PlaySound returns after the sound event completes.
Public Const SND_ASYNC = &H1            ' The sound is played asynchronously and PlaySound returns immediately after beginning the sound. To terminate an asynchronously played waveform sound, call PlaySound with pszSound set to NULL.
Public Const SND_NODEFAULT = &H2        ' No default sound event is used. If the sound cannot be found, PlaySound returns silently without playing the default sound.
Public Const SND_MEMORY = &H4           ' A sound event's file is loaded in RAM. The parameter specified by pszSound must point to an image of a sound in memory.
Public Const SND_LOOP = &H8             ' The sound plays repeatedly until PlaySound is called again with the pszSound parameter set to NULL. You must also specify the SND_ASYNC flag to indicate an asynchronous sound event.
Public Const SND_NOSTOP = &H10          ' The specified sound event will yield to another sound event that is already playing. If a sound cannot be played because the resource needed to generate that sound is busy playing another sound, the function immediately returns FALSE without playing the requested sound.  If this flag is not specified, PlaySound attempts to stop the currently playing sound so that the device can be used to play the new sound.
Public Const SND_PURGE = &H40           ' Sounds are to be stopped for the calling task. If pszSound is not NULL, all instances of the specified sound are stopped. If pszSound is NULL, all sounds that are playing on behalf of the calling task are stopped.
Public Const SND_APPLICATION = &H80     ' The sound is played using an application-specific association.
Public Const SND_NOWAIT = &H2000        ' If the driver is busy, return immediately without playing the sound.
Public Const SND_ALIAS = &H10000        ' The pszSound parameter is a system-event alias in the registry or the WIN.INI file. Do not use with either SND_FILENAME or SND_RESOURCE.
Public Const SND_FILENAME = &H20000     ' The pszSound parameter is a filename.
Public Const SND_RESOURCE = &H40004     ' The pszSound parameter is a resource identifier; hmod must identify the instance that contains the resource.
Public Const SND_ALIAS_ID = &H110000    ' The pszSound parameter is a predefined sound identifier.

'***********************************************'
'*              Win32 API Declares             *'
'***********************************************'
Public Type RECT        ' The RECT structure defines the coordinates
    Left As Long        ' of the upper-left and lower-right corners
    Top As Long         ' of a rectangle.
    Right As Long
    Bottom As Long
End Type

Public Type Size        ' The SIZE structure specifies the width and
    cx As Long          ' height of a rectangle.
    cy As Long
End Type

Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef pColorRef As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'***********************************************'
'* TranslateColor                              *'
'*                                             *'
'* This function takes a VB color constant and *'
'* translates it into its RGB equivalent.      *'
'***********************************************'
Public Function TranslateColor(ByVal clrColor As OLE_COLOR, _
                Optional hPalette As Long = 0) As Long

If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
    TranslateColor = CLR_INVALID
End If

End Function

