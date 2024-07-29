Imports System.Drawing.Text
Imports System.Runtime.InteropServices
Public Class CustomFontLoader
    <DllImport("gdi32.dll")>
    Private Shared Function AddFontMemResourceEx(
      ByVal pbFont As IntPtr,
      ByVal cbFont As Integer,
      ByVal pdv As IntPtr,
      ByRef pcFonts As UInteger) As IntPtr
    End Function

    Public Shared Function LoadFont(ByVal fontResourceName As String) As PrivateFontCollection
        Dim fontCollection As New PrivateFontCollection()

        ' Get the resource stream
        Dim fontStream As IO.Stream = GetType(CustomFontLoader).Assembly.GetManifestResourceStream(fontResourceName)
        If fontStream Is Nothing Then
            Throw New Exception($"Resource '{fontResourceName}' not found.")
        End If

        ' Read the font data into a byte array
        Dim fontData(fontStream.Length - 1) As Byte
        fontStream.Read(fontData, 0, fontData.Length)
        fontStream.Close()

        ' Allocate memory and copy the byte array
        Dim fontPtr As IntPtr = Marshal.AllocCoTaskMem(fontData.Length)
        Marshal.Copy(fontData, 0, fontPtr, fontData.Length)

        ' Add the font to the PrivateFontCollection
        Dim pcFonts As UInteger = 0
        AddFontMemResourceEx(fontPtr, fontData.Length, IntPtr.Zero, pcFonts)
        fontCollection.AddMemoryFont(fontPtr, fontData.Length)

        ' Free the memory
        Marshal.FreeCoTaskMem(fontPtr)

        Return fontCollection
    End Function
End Class
