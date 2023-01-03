# SkiaSharpCOM

Use COM Interop to expose the [SkiaSharp](https://github.com/mono/SkiaSharp) library as a COM object that can be accessed from VB6.

Work In Progress

Usage in VB6

```
Private Sub DrawCircle()
    Dim renderer As SkiaRenderer
    Dim picture As IPictureDisp
    Set renderer = New SkiaRenderer

    renderer.CreateBitmap 200, 200
    renderer.DrawCircle 100, 100, 50, 0, 20, True
    Set picture = renderer.ToPicture()
    Picture1.picture = picture
End Sub
```

![image](https://user-images.githubusercontent.com/60496134/210395999-9787be67-3b61-4568-a984-3a649c10606a.png)
