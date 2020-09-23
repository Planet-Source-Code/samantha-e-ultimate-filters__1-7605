Attribute VB_Name = "ImgMod"
Global ImageArray(-4 To 2, -4 To 700, -4 To 700) As Integer
Global X, Y As Integer

Sub loading(i, j)
frmWaiting.Show
frmWaiting.Refresh
    For i = 0 To Y - 1
        For j = 0 To X - 1
            Pixel& = frmFilters.Picture1.Point(j, i)
            Red = Pixel& Mod 256
            Green = ((Pixel& And &HFF00) / 256&) Mod 256&
            Blue = (Pixel& And &HFF0000) / 65536
            ImageArray(0, i, j) = Red
            ImageArray(1, i, j) = Green
            ImageArray(2, i, j) = Blue
        Next
        frmWaiting.ProgressBar1.Value = i * 100 / (Y - 1)
    Next
frmWaiting.Hide
End Sub

