# Custom Font in VB.NET


CustomFontLoader.vb allows you to find and load Fonts imported into the project via Resources.
![devenv_caJyoyscmw](https://github.com/user-attachments/assets/9a809e30-9ce0-43fb-8616-fab8e4d5bea4)
> [!IMPORTANT]
then load the font(s) into the project's form load

 **Imports System.Drawing.Text**

**Public Class Form1**
**Private customFont1 As Font**

**Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load**
**Dim fontCollection1 As PrivateFontCollection = CustomFontLoader.LoadFont("YourAppName.YourFontName.otf") 'You can use .otf or ttf file**
**customFont1 = New Font(fontCollection1.Families(0), 16.2F, FontStyle.Bold)**
**Label1.Font = customFont1**
