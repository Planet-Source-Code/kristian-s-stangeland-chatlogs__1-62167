Attribute VB_Name = "modOpenGL"
Option Explicit

' ================================================================================
' Copyright © 2003-2005 Peter Wilson (peter@sourcecode.net.au)
' ================================================================================

' The GetGlyphOutline function retrieves the outline or bitmap for a character in the TrueType font that is selected into the specified device context.
Private Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hDC&, ByVal uChar&, ByVal fuFormat&, lpgm As GLYPHMETRICS, ByVal cbBuffer&, lpBuffer As Any, lpmat2 As MAT2) As Long

Private lpgm As GLYPHMETRICS

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type FIXED
    fract As Integer
    Value As Integer
End Type

Private Type POINTFX
    X As FIXED
    Y As FIXED
End Type

Private Type TTPOLYCURVE
    wType As Integer
    cpfx As Integer
    apfx As POINTFX
End Type

Private Type TTPOLYGONHEADER
    cb As Long
    dwType As Long
    pfxStart As POINTFX
End Type


' The 'GetGlyphOutline' function retrieves the curve data points in
' the rasterizer's native format and uses the font's design units.
Private Const GGO_NATIVE = 2&

' The GLYPHMETRICS structure contains information about the
' placement and orientation of a glyph in a character cell.
Private Type GLYPHMETRICS
    gmBlackBoxX As Long
    gmBlackBoxY As Long
    gmptGlyphOrigin As POINTAPI
    gmCellIncX As Integer
    gmCellIncY As Integer
End Type

' The MAT2 structure contains the values for a transformation
' matrix used by the GetGlyphOutline function.
Private Type MAT2
    eM11 As FIXED
    eM12 As FIXED
    eM21 As FIXED
    eM22 As FIXED
End Type

' Font Points.
Private Type mdrFontPoint
    X As Double
    Y As Double
    Style As Integer
End Type

' Polygons.
Private Type mdrPolygon
    Vertex() As mdrFontPoint
End Type

Private Const GDI_ERROR As Long = &HFFFF
Private Const TT_PRIM_LINE = 1       ' Curve is a polyline.
Private Const TT_PRIM_QSPLINE = 2    ' Curve is a quadratic Bézier spline.
Private Const TT_PRIM_CSPLINE = 3    ' Curve is a cubic Bézier spline.

Private m_intPointCount As Integer
Private m_objFontPoints() As mdrFontPoint
Private m_objPolygons() As mdrPolygon

Private Const g_sngPIDivideBy180 As Single = 0.0174533!

Public Function CreateFonts(hDC As Long) As GLuint

    Dim listID As GLuint
    
    listID = glGenLists(24)
    
    If listID = 0 Then
        ' Inform about the error
        Err.Raise Err.LastDllError, "CreateFonts", "Could not create a new list identifer."
    End If
    
    CreateFonts = listID
    
    ' 24 characters A-Z
    Dim intN As Integer
    
    For intN = 33 To 126
    
        glNewList listID + (intN - 33), lstCompile
        
        GetGlyphs hDC, 0.5, intN
        DrawGlyph
            
        ' Draw a space.
        glTranslated lpgm.gmCellIncX, 0, 0
        
        glEndList
    
    Next
    
End Function

Public Sub DrawGlyph()
    
    Dim intPoly As Integer
    Dim intVertex As Integer
    Dim sngX As Single
    Dim sngY As Single
       
    For intPoly = LBound(m_objPolygons) To UBound(m_objPolygons)
        
        Call glBegin(bmLineLoop)
        For intVertex = LBound(m_objPolygons(intPoly).Vertex) To UBound(m_objPolygons(intPoly).Vertex)
            
            sngX = m_objPolygons(intPoly).Vertex(intVertex).X
            sngY = m_objPolygons(intPoly).Vertex(intVertex).Y
            
            Call glVertex3f(sngX, sngY, 0)
            
        Next intVertex
        Call glEnd
        
    Next intPoly
    
End Sub

Private Function DoubleFromFixed(f As FIXED) As Double
    
   Dim d As Double
   
   d = CDbl(f.Value)
   
   If f.fract < 0 Then
      d = d + (32768 + (f.fract And &H7FFF)) / 65536#
   Else
      d = d + CDbl(f.fract) / 65536#
   End If
   
   DoubleFromFixed = d
   
End Function

Private Function FixedFromDouble(ByVal d As Double) As FIXED

   Dim f As FIXED
   Dim i As Long
   
   ' Calculate the Value portion
   ' Note: -1.2 must be rounded to -2,  The Value portion can be
   ' positive or negative, but the Fract portion can only be
   ' positive.  Hence -1.2 is stored as -2 + 0.8
   
   i = Int(d)
   
   If i < 0 Then
      f.Value = &H8000 Or CInt(i And &H7FFF)
   Else
      f.Value = CInt(i And &H7FFF)
   End If
      
   i = (CLng(d * 65536#) And &HFFFF&)
   
   If (i And &H8000&) = &H8000& Then
      f.fract = &H8000 Or CInt(i And &H7FFF&)
   Else
      f.fract = CInt(i And &H7FFF&)
   End If
   
   FixedFromDouble = f
   
End Function

Private Sub Spline(Resolution As Single, PointA As mdrFontPoint, PointB As mdrFontPoint, PointC As mdrFontPoint)
    
    ' Interpolates a curved surface give some control points.
    ' Resolution settings:
    '   1.0     Lowest value that will approximate a curve.
    '   0.5     Better
    '   0.25    Good
    '  >0.1     Best
    
    Dim sngT As Single
    Dim sngX As Single
    Dim sngY As Single
    
    sngT = 0
    
    While Round(sngT, 6) <= 1
    
        sngX = (PointA.X - 2 * PointB.X + PointC.X) * sngT ^ 2 + (2 * PointB.X - 2 * PointA.X) * sngT + PointA.X
        sngY = (PointA.Y - 2 * PointB.Y + PointC.Y) * sngT ^ 2 + (2 * PointB.Y - 2 * PointA.Y) * sngT + PointA.Y
        m_intPointCount = m_intPointCount + 1
        
        ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
        
        m_objFontPoints(m_intPointCount).Style = TT_PRIM_QSPLINE
        m_objFontPoints(m_intPointCount).X = sngX
        m_objFontPoints(m_intPointCount).Y = sngY
        
        sngT = sngT + Resolution
    Wend
    
End Sub

Public Sub GetGlyphs(ByVal hDC&, Resolution As Single, WhichByte As Integer)
    
    Dim lngTotalNativeBuffer As Long
    Dim lpmat2 As MAT2
    Dim abytBuffer() As Byte
    Dim intN As Integer
    Dim objPolyHeader() As TTPOLYGONHEADER
    Dim objPolyCurve() As TTPOLYCURVE
    Dim objPointFX() As POINTFX
    Dim lngTotalPolygonHeader As Long
    Dim lngIndex As Long
    Dim lngIndexPolygonHeader As Long
    Dim lngPolygonCount As Long
    
    lpmat2 = MatrixRotationZ(0)

    ' Get the required buffer size  (This works)
    ' The function retrieves the curve data points in the rasterizer's native format and uses the font's design units.
    lngTotalNativeBuffer = GetGlyphOutline(hDC, WhichByte, GGO_NATIVE, lpgm, 0, ByVal 0&, lpmat2)
    If lngTotalNativeBuffer <> GDI_ERROR Then

        ' Set the buffer size
        ReDim abytBuffer(lngTotalNativeBuffer - 1) As Byte
        
        ' Then retrieve the information
        If GetGlyphOutline(hDC, WhichByte, GGO_NATIVE, lpgm, lngTotalNativeBuffer, abytBuffer(0), lpmat2) <> GDI_ERROR Then

            ReDim objPolyHeader(0) As TTPOLYGONHEADER
            ReDim objPolyCurve(0) As TTPOLYCURVE

            lngIndex = 0
            m_intPointCount = 0
            
            Do
                ' ==================================================================
                ' Copy a PolygonHeader into memory (and increment the buffer index).
                ' ==================================================================
                CopyMemory objPolyHeader(0), abytBuffer(lngIndex), 16
                lngTotalPolygonHeader = objPolyHeader(0).cb
                lngIndex = lngIndex + 16
                lngIndexPolygonHeader = 16
                
                ' ===================================================================================
                ' A PolygonHeader has a start point, that can either be the start of a straight line,
                ' or the start of a quadratic Bézier spline. i.e. Point A, in an [A,B,C] curve.
                ' ===================================================================================
                m_intPointCount = m_intPointCount + 1
                ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                
                m_objFontPoints(m_intPointCount).Style = 0 ' ie. Starting point.
                m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPolyHeader(0).pfxStart.X)
                m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPolyHeader(0).pfxStart.Y)
                               
                Do
                    ' ======================================================
                    ' At least one PolyCurve always follows a PolygonHeader.
                    ' PolyCurve always has a least one starting PointFX.
                    ' ======================================================
                    CopyMemory objPolyCurve(0), abytBuffer(lngIndex), 12
                    lngIndex = lngIndex + 12
                    lngIndexPolygonHeader = lngIndexPolygonHeader + 12
                    
                    ' =======================================
                    ' Load additional PointFX values (if any).
                    ' =======================================
                    If objPolyCurve(0).cpfx > 1 Then
                        ReDim objPointFX((objPolyCurve(0).cpfx - 2)) As POINTFX
                        CopyMemory objPointFX(0), abytBuffer(lngIndex), (8 * (objPolyCurve(0).cpfx - 1))
                        lngIndex = lngIndex + (8 * (objPolyCurve(0).cpfx - 1))
                        lngIndexPolygonHeader = lngIndexPolygonHeader + (8 * (objPolyCurve(0).cpfx - 1))
                    End If
                
                    ' Part A) Create the initial polycurve point...
                    m_intPointCount = m_intPointCount + 1
                    ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                                   m_objFontPoints(m_intPointCount).Style = TT_PRIM_LINE
                                   m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPolyCurve(0).apfx.X)
                                   m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPolyCurve(0).apfx.Y)
                
                    ' Post-Process points depending on whether they are straight lines, or curves.
                    ' ===========================================================================
                    If objPolyCurve(0).wType = TT_PRIM_LINE Then
    
                        ' ============================
                        ' PointFX(0..n) is a polyline.
                        ' ============================
                        
                        ' Part B) ...Create subsequent points.
                        If objPolyCurve(0).cpfx > 1 Then
                            For intN = LBound(objPointFX) To UBound(objPointFX)
    
                                m_intPointCount = m_intPointCount + 1
                                ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                                               m_objFontPoints(m_intPointCount).Style = TT_PRIM_LINE
                                               m_objFontPoints(m_intPointCount).X = DoubleFromFixed(objPointFX(intN).X)
                                               m_objFontPoints(m_intPointCount).Y = DoubleFromFixed(objPointFX(intN).Y)
                            Next intN
                        End If
    
                    ElseIf objPolyCurve(0).wType = TT_PRIM_QSPLINE Then
                    
                        ' ======================================================================
                        ' PointFX(0..n) is a quadratic Bézier spline.
                        ' Load the Spline's Control points first, then create the Spline itself.
                        ' ======================================================================
                        Dim lngControlPoint  As Long
                        Dim objSplinePoints() As mdrFontPoint
                        ReDim objSplinePoints(1) As mdrFontPoint
                        
                        ' The last defined point is Point A (from a previous object/header).
                        ' ==================================================================
                        lngControlPoint = 0
                        objSplinePoints(0).Style = -1 ' ie. Control Point for spline.
                        objSplinePoints(0).X = m_objFontPoints(m_intPointCount - 1).X
                        objSplinePoints(0).Y = m_objFontPoints(m_intPointCount - 1).Y
                        
                        ' This is the first control point B.
                        lngControlPoint = lngControlPoint + 1
                        objSplinePoints(1).Style = -1 ' ie. Control Point for spline.
                        objSplinePoints(1).X = m_objFontPoints(m_intPointCount).X
                        objSplinePoints(1).Y = m_objFontPoints(m_intPointCount).Y
                        
                        If objPolyCurve(0).cpfx > 1 Then ' Splines always have a number greater than 1.
                            For intN = LBound(objPointFX) To UBound(objPointFX)
                            
                                lngControlPoint = lngControlPoint + 1
                                ReDim Preserve objSplinePoints(lngControlPoint) As mdrFontPoint
                                               objSplinePoints(lngControlPoint).Style = -1 ' ie. Control Point for spline.
                                               objSplinePoints(lngControlPoint).X = DoubleFromFixed(objPointFX(intN).X)
                                               objSplinePoints(lngControlPoint).Y = DoubleFromFixed(objPointFX(intN).Y)
                            Next intN
                        End If
                        
                        If Resolution > 0 Then ' Do Curved surfaces...
                        
                            ' ========================================================================================
                            ' At this point, we have an array of control points that help make up a series of splines.
                            ' ie. objSplinePoints(0 to lngControlPoint)
                            ' Note: They may be in the form A,B,B,B,C.
                            '       In which case, new C's need to be found inbetween the B's (except the last one).
                            ' ========================================================================================
                            Dim intB As Integer
                            Dim PointA As mdrFontPoint
                            Dim PointB As mdrFontPoint
                            Dim PointC As mdrFontPoint
                            
                            intB = 1
                            PointC = objSplinePoints(intB - 1)
                            
                            lngControlPoint = lngControlPoint + 1
                            ReDim Preserve m_objFontPoints(m_intPointCount) As mdrFontPoint
                            
                            m_objFontPoints(m_intPointCount).Style = -1 ' ie. Control Point for spline.
                            m_objFontPoints(m_intPointCount).X = PointC.X
                            m_objFontPoints(m_intPointCount).Y = PointC.Y
                            
                            Do
                                PointA = PointC
                                PointB = objSplinePoints(intB)
                                
                                If intB < (lngControlPoint - 2) Then
                                    ' Find a new midpoint
                                    PointC.X = (objSplinePoints(intB).X + objSplinePoints(intB + 1).X) / 2
                                    PointC.Y = (objSplinePoints(intB).Y + objSplinePoints(intB + 1).Y) / 2
                                Else
                                    PointC = objSplinePoints(intB + 1)
                                End If
                                
                                Call Spline(Resolution, PointA, PointB, PointC)
                                intB = intB + 1
                                
                            Loop Until intB = (lngControlPoint - 1)
                            
                        End If ' Ignore curved surfaces if m_sngResolution = 0.
                        
                    Else
                        ' Unsupport style of font.
                        Err.Raise vbObjectError + 1001, "GetGlyphs", "Unsupported Font Style. Please contact technical support (http://dev.midar.com) for additional information on this error."
                    End If
                
                Loop Until lngIndexPolygonHeader = lngTotalPolygonHeader
            Loop Until lngIndex = lngTotalNativeBuffer


            ' =====================================================================
            ' Clean up, Remove Duplicates, and separate into Polygons and Vertices.
            ' =====================================================================
            Dim intVertexCount As Integer
            Dim intPolyCount As Integer
            intPolyCount = -1
            For intN = LBound(m_objFontPoints) + 1 To UBound(m_objFontPoints)
                If m_objFontPoints(intN).Style = 0 Then
                    intVertexCount = -1
                    intPolyCount = intPolyCount + 1
                    ReDim Preserve m_objPolygons(intPolyCount)
                End If
                
                If (m_objFontPoints(intN - 1).X = m_objFontPoints(intN).X) And _
                   (m_objFontPoints(intN - 1).Y = m_objFontPoints(intN).Y) Then
                    ' Do nothing because this is a duplicate item.
                Else
                    ' Add new vertex.
                    intVertexCount = intVertexCount + 1
                    ReDim Preserve m_objPolygons(intPolyCount).Vertex(intVertexCount)
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).Style = m_objFontPoints(intN).Style
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).X = m_objFontPoints(intN).X
                                   m_objPolygons(intPolyCount).Vertex(intVertexCount).Y = m_objFontPoints(intN).Y
                End If
                
            Next intN


        Else
            Err.Raise Err.LastDllError
        End If
    Else
        Err.Raise Err.LastDllError
    End If
    
End Sub

Private Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Private Function MatrixIdentity() As MAT2

    ' The identity matrix is used as the starting point for matrices
    ' that will modify vertex values to create rotations, translations,
    ' and any other transformations that can be represented by a 2x2 matrix.
    '
    ' Notice that...
    '   * the 1's go diagonally down?
    '   * rc stands for Row Column. Therefore, rc12 means Row1, Column 2.
    
    With MatrixIdentity
        ' Value Part
        .eM11.Value = 1: .eM12.Value = 0
        .eM21.Value = 0: .eM22.Value = 1
        
        ' Fraction Part
        .eM11.fract = 0: .eM12.fract = 0
        .eM21.fract = 0: .eM22.fract = 0
    End With
    
End Function

Private Function MatrixRotationZ(Radians As Single) As MAT2
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 2)
    sngSine = Round(Sin(Radians), 2)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity()
    
    ' Z-Axis rotation.
    With MatrixRotationZ
        ' Actual Rotation values.
        .eM11 = FixedFromDouble(sngCosine)
        .eM21 = FixedFromDouble(sngSine)
        .eM12 = FixedFromDouble(-sngSine)
        .eM22 = FixedFromDouble(sngCosine)
        
        ' Increase Resolution by multiplying the matrix with 2048.
        ' (2048 is a size recommended my www.microsoft.com/typography/)
        .eM11 = FixedFromDouble(DoubleFromFixed(.eM11) * 2048)
        .eM21 = FixedFromDouble(DoubleFromFixed(.eM21) * 2048)
        .eM12 = FixedFromDouble(DoubleFromFixed(.eM12) * 2048)
        .eM22 = FixedFromDouble(DoubleFromFixed(.eM22) * 2048)
        
    End With

End Function


