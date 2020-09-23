VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStarWars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ============================================================
' Copyright © 2003-2005 Peter Wilson (peter@sourcecode.net.au)
' Modified by Kristian S. Stangeland (kr_stang@hotmail.com)
' ============================================================

' Used to play MIDI-files
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' Events used to communicate with the parent control
Event AnimationEnded()
Event ReachedEnd()

Private Type Line
    Text As String
    Color As Long
    ScaleX As Double
    ScaleY As Double
    LineWith As Long
End Type

Private Type Star
    X As Double
    Y As Double
    Color As Double
End Type

' Public varaibles to change
Public ScrollingSpeed As Double
Public FramesPerSecond As Long
Public WindowWidth As Long
Public WindowHeight As Long

' Positions used in the drawing-process
Private dblD As Double
Private dblA As Double
Private dblB As Double

' All lines and starts we're going to render
Private Lines() As Line
Private Stars() As Star
Private LineCount As Long

' The rendering contex
Private m_hglrc As Long

' Used to stop the loop
Private bEnd As Boolean

' Display List Base
Private DisplayListBase As GLuint

Public Sub AllocateStars(Amount As Long)
    
    Dim Tell As Long
    
    ' Reallocate the array
    ReDim Stars(Amount)

    ' Initialize the random generator
    Randomize

    ' Now, just set the points and colors to a random value
    For Tell = 0 To Amount
        
        ' Set the properties of this star
        Stars(Tell).X = (Rnd * (15 * 2)) - 15
        Stars(Tell).Y = (Rnd * (10 * 2)) - 10
        Stars(Tell).Color = Rnd
        
    Next

End Sub

Public Sub AddLine(sText As String, Optional Color = vbYellow, Optional ScaleX = 1, Optional ScaleY = 1, Optional LineWith = 1)

    ' Now, increase the line-count
    LineCount = LineCount + 1
    
    ' Then reallocate the array
    ReDim Preserve Lines(LineCount)

    ' And set the different properties of this line
    Lines(LineCount).Text = sText
    Lines(LineCount).Color = Val(Color)
    Lines(LineCount).ScaleX = Val(ScaleX)
    Lines(LineCount).ScaleY = Val(ScaleY)
    Lines(LineCount).LineWith = Val(LineWith)

End Sub

Public Sub ClearLines()

    ' Firstly, erase the array
    Erase Lines
    
    ' Then reset the line count
    LineCount = -1

End Sub

Public Sub StopAnimation()

    ' Close our MIDI-file
    mciSendString "close starwars", 0&, 0, 0

    ' Set the loop to stop
    bEnd = True

End Sub

Public Sub StartAnimation(hDC As Long, Optional StarAmount As Long = 300)

    ' Load the MIDI-file
    mciSendString "open " & Chr(34) & App.Path & "\..\Data\Music\StarWars.mid" & Chr(34) & " type sequencer alias starwars", 0&, 0, 0
    
    ' Start playing it
    mciSendString "play starwars from 14 notify", 0&, 0, 0
    
    ' Reset variables
    bEnd = False
    dblD = 0
    dblA = 0
    dblB = 0
    
    ' Allocate 50 stars
    AllocateStars StarAmount
    
    ' Initialize the graphics engine (OpenGL)
    InitializeOpenGL hDC
    
    ' Retrive the font to use
    DisplayListBase = CreateFonts(hDC)

    ' Start animation loop
    Do Until bEnd
        
        ' Draw the scene
        DrawScene hDC
        
        ' Perfom events and wait
        DoEvents
        Sleep (1000 / FramesPerSecond)
        
    Loop
    
    ' Delete and release the contex
    wglMakeCurrent hDC, 0&
    wglDeleteContext m_hglrc
    
    ' The animation has ended
    RaiseEvent AnimationEnded

End Sub

Private Sub InitializeOpenGL(hDC As Long)
        
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim iPixelFormat As Long

    ' The PIXELFORMATDESCRIPTOR structure describes the pixel format of a drawing surface.
    ZeroMemory pfd, Len(pfd)
    pfd.nSize = Len(pfd)
    
    With pfd
        .nVersion = 1
        .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
        .iPixelType = PFD_TYPE_RGBA
        .cDepthBits = 32
        .iLayerType = PFD_MAIN_PLANE
    End With
    
    ' The ChoosePixelFormat function attempts to match an appropriate pixel
    iPixelFormat = ChoosePixelFormat(hDC, pfd)
    
    ' The SetPixelFormat function sets the pixel format of the specified device context
    SetPixelFormat hDC, iPixelFormat, pfd
    
    ' Note: Set the pixel format of the device context before creating a rendering context.
    m_hglrc = wglCreateContext(hDC)

    ' The wglMakeCurrent function makes a specified OpenGL rendering context the
    ' calling thread's current rendering context.
    wglMakeCurrent hDC, m_hglrc

End Sub

Private Sub PrintText(Text As String)
    
    Dim dblScale As Double
    Dim varWords As Variant
    Dim intW As Integer
    Dim intN As Integer
    Dim intLen As Integer
    Dim Bytes() As Byte
    
    ' Clean up
    Text = Trim(Text)
    
    ' Only try to write the text if there, in fact, are something to write
    If LenB(Text) Then
    
        ' Split each word
        varWords = Split(Text, " ")
        
        ' Now, set the scale to use
        dblScale = (1 / 4096)
        
        ' Start drawing
        glPushMatrix
            
            ' Scale everything
            glScaled dblScale, dblScale, dblScale
        
            ' Loop through all words
            For intW = LBound(varWords) To UBound(varWords)
            
                ' Get length of text.
                intLen = Len(varWords(intW))
                    
                ' Size array
                ReDim Bytes(intLen - 1) As Byte
                
                glListBase (-33 + DisplayListBase)
                
                For intN = 0 To intLen - 1
                    ' Retrive the byte (as ASCII)
                    Bytes(intN) = Asc(Mid(varWords(intW), intN + 1, 1))
                Next
                
                ' Draw word
                glCallLists intLen, GL_UNSIGNED_BYTE, Bytes(0)
            
                ' Draw Space between words (since there is no character for a space)
                glTranslatef (1 / dblScale) * 3, 0, 0
                
            Next
        
        ' Drawing is finished
        glPopMatrix
        
    End If

End Sub

' Draws the Scence.
Private Sub DrawScene(hDC As Long)
    
    Dim Tell As Long, lRed As Long, lGreen As Long, lBlue As Long
    Dim dblAspectRatio As Double, Pos As Long
    Dim objFogColour(3) As GLfloat
    
    ' Let the user resize the window, and maintain the correct aspect ratio.
    dblAspectRatio = ((WindowWidth - 64) / (WindowHeight - 64))
    
    ' The glClearColor function specifies clear values for the color buffers.
    glClearColor 0, 0, 0, 0

    ' Clear the buffers.
    glClear clrColorBufferBit Or clrDepthBufferBit

    ' The glMatrixMode function specifies which matrix is the current matrix.
    glMatrixMode mmProjection
    glLoadIdentity
    
    ' MatrixViewMapping_Per (This is a nice one to change).
    gluPerspective 90, dblAspectRatio, 0.1, 196
    
    ' ViewPort
    glViewport 0, 0, WindowWidth, WindowHeight
    
    ' Enable Fog.
    glFogf fogMode, GL_LINEAR
    glFogf fogStart, 32
    glFogf fogEnd, 196
    glFogf fogDensity, 1
    
    glFogfv fogColor, objFogColour(0)
    glEnable glcFog
    
    glMatrixMode (mmModelView)
    glLoadIdentity

    ' Draw the stars
    For Tell = LBound(Stars) To UBound(Stars)
    
        ' Set the color
        glColor3f Stars(Tell).Color, Stars(Tell).Color, Stars(Tell).Color
        
        ' Then start drawing the star
        glBegin bmPoints
            glVertex3f Stars(Tell).X, Stars(Tell).Y, -10
        glEnd
    
    Next
        
    ' View Orientation - the speed
    dblD = dblD - ScrollingSpeed
    gluLookAt 93, dblD, 93, 93, dblD + 64, 0, 0, 1, 0

    ' Move downwards a bit
    glTranslated 0, -30, 0
    
    ' Loop through all lines
    For Tell = 0 To LineCount
    
        ' Now, start drawing
        glPushMatrix
        
        ' Set the scale to use
        glScaled Lines(Tell).ScaleX, Lines(Tell).ScaleY, 0
        
        ' Then, set the line with
        glLineWidth Lines(Tell).LineWith
        
        ' Retrive the color
        SplitRGB Lines(Tell).Color, lRed, lGreen, lBlue
        
        ' And set the color
        glColor3f lRed / 255, lGreen / 255, lBlue / 255
        
        ' Set the position to draw to
        glTranslated 0, Pos, 0
        
        ' And finally, print the text
        PrintText Lines(Tell).Text
        
        ' Now, decrease the position
        Pos = Pos - 7
        
        ' We're finsihed drawing
        glPopMatrix
        
    Next

    ' Update display.
    glFlush
    SwapBuffers hDC
    
    ' And see if we've reached the end
    If dblD < -685 Then
        RaiseEvent ReachedEnd
    End If
    
End Sub

Private Sub SplitRGB(ByVal lColor As Long, lRed As Long, lGreen As Long, lBlue As Long)
    
    ' Return the red, green and blue part of the RGB-value
    lRed = lColor And &HFF
    lGreen = (lColor And &HFF00&) \ &H100&
    lBlue = (lColor And &HFF0000) \ &H10000
    
End Sub

Private Sub Class_Initialize()

    ' Set default values of the array
    ClearLines

    ' Now, set the remaining variables
    ScrollingSpeed = 0.1
    FramesPerSecond = 60
    WindowWidth = 640
    WindowHeight = 480

End Sub

Private Sub Class_Terminate()

    ' Stop the animation just in case
    StopAnimation

End Sub
