Option Strict On
Option Explicit On
Imports System.Drawing.Imaging

Public Class PerPixelAlphaForm
    Inherits Form

    'Dim CurrBitmap As Bitmap

    Public Sub New()
        FormBorderStyle = Windows.Forms.FormBorderStyle.None
    End Sub

    Public Sub SetBitmap(ByVal bitmap As Bitmap)
        SetBitmap(bitmap, 255)
    End Sub

    Public Sub SetBitmap(ByVal bitmap As Bitmap, ByVal opacity As Byte)
        If Not (bitmap.PixelFormat = PixelFormat.Format32bppArgb) Then
            Throw New ApplicationException("The bitmap must be 32ppp with alpha-channel.")
        End If
        'CurrBitmap = bitmap
        Dim screenDc As IntPtr = Win32.GetDC(IntPtr.Zero)
        Dim memDc As IntPtr = Win32.CreateCompatibleDC(screenDc)
        Dim hBitmap As IntPtr = IntPtr.Zero
        Dim oldBitmap As IntPtr = IntPtr.Zero
        Try
            hBitmap = bitmap.GetHbitmap(Color.FromArgb(0))
            oldBitmap = Win32.SelectObject(memDc, hBitmap)
            Dim size As Win32.Size = New Win32.Size(bitmap.Width, bitmap.Height)
            Dim pointSource As Win32.Point = New Win32.Point(0, 0)
            Dim topPos As Win32.Point = New Win32.Point(Left, Top)
            Dim blend As Win32.BLENDFUNCTION = New Win32.BLENDFUNCTION
            blend.BlendOp = Win32.AC_SRC_OVER
            blend.BlendFlags = 0
            blend.SourceConstantAlpha = opacity
            blend.AlphaFormat = Win32.AC_SRC_ALPHA
            Win32.UpdateLayeredWindow(Handle, screenDc, topPos, size, memDc, pointSource, 0, blend, Win32.ULW_ALPHA)
        Finally
            Win32.ReleaseDC(IntPtr.Zero, screenDc)
            If Not (hBitmap = IntPtr.Zero) Then
                Win32.SelectObject(memDc, oldBitmap)
                Win32.DeleteObject(hBitmap)
            End If
            Win32.DeleteDC(memDc)
        End Try
    End Sub

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or (524288)
            Return cp
        End Get
    End Property

    Private Sub PerPixelAlphaForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'PerPixelAlphaForm
        '
        Me.ClientSize = New System.Drawing.Size(284, 264)
        Me.Name = "PerPixelAlphaForm"
        Me.ResumeLayout(False)

    End Sub
End Class

Class Win32

    Public Enum Bool
        [False] = 0
        [True]
    End Enum

    '<StructLayout(LayoutKind.Sequential)> _
    Public Structure Point
        Public x As Int32
        Public y As Int32

        Public Sub New(ByVal x As Int32, ByVal y As Int32)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    '<StructLayout(LayoutKind.Sequential)> _
    Public Structure Size
        Public cx As Int32
        Public cy As Int32

        Public Sub New(ByVal cx As Int32, ByVal cy As Int32)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    '<StructLayout(LayoutKind.SequentialPack = 1)> _
    Structure ARGB
        Public Blue As Byte
        Public Green As Byte
        Public Red As Byte
        Public Alpha As Byte
    End Structure

    '<StructLayout(LayoutKind.Sequential, , 1)> _
    Public Structure BLENDFUNCTION
        Public BlendOp As Byte
        Public BlendFlags As Byte
        Public SourceConstantAlpha As Byte
        Public AlphaFormat As Byte
    End Structure
    Public Const ULW_COLORKEY As Int32 = 1
    Public Const ULW_ALPHA As Int32 = 2
    Public Const ULW_OPAQUE As Int32 = 4
    Public Const AC_SRC_OVER As Byte = 0
    Public Const AC_SRC_ALPHA As Byte = 1

    <DllImport("user32.dll")> _
    Public Shared Function UpdateLayeredWindow(ByVal hwnd As IntPtr, ByVal hdcDst As IntPtr, ByRef pptDst As Point, ByRef psize As Size, ByVal hdcSrc As IntPtr, ByRef pprSrc As Point, ByVal crKey As Int32, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Int32) As Bool
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function GetDC(ByVal hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Public Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    End Function

    <DllImport("gdi32.dll")> _
    Public Shared Function CreateCompatibleDC(ByVal hDC As IntPtr) As IntPtr
    End Function

    <DllImport("gdi32.dll")> _
    Public Shared Function DeleteDC(ByVal hdc As IntPtr) As Bool
    End Function

    <DllImport("gdi32.dll")> _
    Public Shared Function SelectObject(ByVal hDC As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function

    <DllImport("gdi32.dll")> _
    Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Bool
    End Function
End Class

