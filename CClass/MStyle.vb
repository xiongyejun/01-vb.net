Module MStyle    '
    'Windows API calls to do all the dirty work
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
    Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Integer) As Integer

    'Window API constants
    Private Const GWL_STYLE As Integer = (-16)           'The offset of a window's style
    Private Const SC_CLOSE As Integer = &HF060           'Close menu item

    Sub SetUserformAppearance(ByVal hwnd As Integer)
        Dim lStyle As Integer
        Dim hMenu As Integer

        'Get the normal windows style bits
        lStyle = GetWindowLong(hwnd, GWL_STYLE)

        'The Close button is handled by removing it from the  control menu, not through a window style bit     
        'We don't want it, so delete it from the control menu
        hMenu = GetSystemMenu(hwnd, 0)
        DeleteMenu(hMenu, SC_CLOSE, 0&)

        'Refresh the window with the changes
        DrawMenuBar(hwnd)
    End Sub

End Module
