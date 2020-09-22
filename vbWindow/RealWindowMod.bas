Attribute VB_Name = "RealWindowMod"
' This is a module to create a window from scratch in Visual Basic.
' I always wanted to do this ever sence i left VB for C++ 2 years back.
' You can do anything you want to this window like you would in Win32 C++.
' The sky is the limit.

' Author: Andrew Heinlein [Mouse]
' Web: www.mouseindustries.com
' Email: mouse@mouseindustries.com

' WARNING: If your brave and decide to debug this, besure to save your work
' before doing so.  You are now `subclassing` and Visual Basic wasnt meant to subclass.
' Also, if you get the "Failed to register window" message, just change the class
' name in the AppMain.BAS

' Added on May 15th, 2002
' how to impliment common controls from scratch

Option Explicit

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(32) As Byte
End Type

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const CS_PARENTDC = &H80
Private Const WS_OVERLAPPEDWINDOW = &HCF0000
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const IDC_ARROW = &H7F00
Private Const COLOR_WINDOW = &H5
Private Const SW_SHOW = &H5
Private Const WM_DESTROY = &H2
Private Const WM_PAINT = &HF
Private Const WM_CREATE = &H1
Private Const DT_CENTER = &H1
Private Const CW_USEDEFAULT = &H80000000
Private Const WS_CHILD = &H40000000
Private Const SS_SUNKEN = &H1000
Private Const SS_CENTER = &H1
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_CLIENTEDGE = &H200
Private Const LBS_NOTIFY = &H1
Private Const WS_VSCROLL = &H200000
Private Const LB_ADDSTRING = &H180
Private Const WM_COMMAND = &H111
Private Const CBS_DROPDOWNLIST = &H3
Private Const CBS_AUTOHSCROLL = &H40
Private Const CBS_HASSTRINGS = &H200
Private Const CB_ADDSTRING = &H143
Private Const CBS_DISABLENOSCROLL = &H800&
Private Const CB_SETCURSEL = &H14E
Private Const LBN_SELCHANGE = &H1
Private Const LB_GETTEXT = &H189
Private Const LB_GETCURSEL = &H188
Private Const WM_TIMER = &H113

' dims for the dynamic controls
Dim hWndButton As Long
Private Const IDC_BUTTON = &H1000
Dim hWndEditBox As Long
Private Const IDC_EDIT = &H1001
Dim hWndStatic As Long
Private Const IDC_STATIC = &H1002
Dim hWndList As Long
Private Const IDC_LIST = &H1003
Dim hWndCombo As Long
Private Const IDC_COMBO = &H1004
Dim hWndStaticTimer As Long
Private Const IDC_STATIC_TIMER = &H1005
Private Const ID_TIMER = &HCAFEBABE

Private Function MainWndProc(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' this should be a Select Case setup, i didnt
    ' set it up as so since its easier to read as IF statements


    ' i put this here so that you can see how to process a message.
    ' this is the WM_PAINT message where it repaints the window.
    ' lets put "Hello World!" at the top of it like they do on
    ' the win32 C++ pre-made projects
    If message = WM_PAINT Then
        Dim rt As RECT
        Dim hdc As Long
        Dim ps As PAINTSTRUCT
        
        GetClientRect hwnd, rt
        
        hdc = BeginPaint(hwnd, ps)
        DrawText hdc, "Hello World!", Len("Hello World!"), rt, DT_CENTER
        EndPaint hwnd, ps
        
        ' since we handled this message, return 0. dont let the
        ' DefWindowProc handle it
        MainWndProc = 0
        Exit Function
    End If

    ' watch for WM_DESTROY message, if its sent, then let the GetMessage loop in
    ' CreateNewWindow know so it breaks out of the GetMessage loop
    If message = WM_DESTROY Then
        KillTimer hwnd, ID_TIMER
        PostQuitMessage 0
        MainWndProc = 0
        Exit Function
    End If
    
    ' capture the CREATE message and create our controls dynamically
    If message = WM_CREATE Then
        CreateControls hwnd
        MainWndProc = 0
        Exit Function
    End If
    
    ' capture the COMMAND message to process control commands
    If message = WM_COMMAND Then
        Dim wmId As Integer, wmEvent As Integer
        wmId = LOWORD(wParam)
        wmEvent = HIWORD(wParam)
        
        ' did user press the button?
        If wmId = IDC_BUTTON Then
            MessageBox hwnd, "Hello button!", "dynamic button", 0
            
            MainWndProc = 0
            Exit Function
        End If
        
        ' was the listbox clicked on?
        If wmId = IDC_LIST Then
            If wmEvent = LBN_SELCHANGE Then
                ' create a buffer
                Dim listBoxText As String
                listBoxText = Space(255)
                ' get the selected item's text
                SendMessage lParam, LB_GETTEXT, SendMessage(lParam, LB_GETCURSEL, 0, 0), listBoxText
                MessageBox hwnd, "Selection changed in list box to: " & listBoxText, "Selection changed", 0
                
                MainWndProc = 0
                Exit Function
            End If
        End If
    End If
    
    ' our timer was set off, lets up our counter window.
    ' Timers are not forgiving. dont put anything in here like a messagebox
    ' because it will set off another if the first wasnt closed out, giving
    ' you one hell of a mess to clean up
    If message = WM_TIMER Then
        Dim winText As String
        winText = Space(255)
        Dim winLen As Long
        winLen = GetWindowText(hWndStaticTimer, winText, 255)
        winText = Left(winText, winLen)
        winText = CLng(winText) + 1
        SetWindowText hWndStaticTimer, winText
        
        MainWndProc = 0
        Exit Function
    End If
    
    MainWndProc = DefWindowProc(hwnd, message, wParam, lParam)
End Function

Private Sub CreateControls(ByVal parent As Long)
    hWndButton = CreateWindowEx(0, "Button", "Real Button in VB", WS_CHILD Or WS_VISIBLE, 10, 20, 140, 30, parent, IDC_BUTTON, App.hInstance, 0)
    hWndStatic = CreateWindowEx(0, "Static", "Real Static Label in VB", WS_CHILD Or WS_VISIBLE, 10, 100, 140, 30, parent, IDC_BUTTON, App.hInstance, 0)
    hWndStaticTimer = CreateWindowEx(0, "Static", "0", WS_CHILD Or WS_VISIBLE Or SS_SUNKEN Or SS_CENTER, 10, 300, 140, 30, parent, IDC_STATIC_TIMER, App.hInstance, 0)
    hWndEditBox = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "Real Editbox in VB", WS_CHILD Or WS_VISIBLE, 10, 60, 140, 25, parent, IDC_EDIT, App.hInstance, 0)
    hWndList = CreateWindowEx(WS_EX_CLIENTEDGE, "listbox", vbNullString, WS_CHILD Or WS_VISIBLE Or LBS_NOTIFY Or WS_VSCROLL, 10, 140, 140, 100, parent, IDC_LIST, App.hInstance, 0)
    hWndCombo = CreateWindowEx(WS_EX_CLIENTEDGE, "combobox", vbNullString, WS_CHILD Or WS_VISIBLE Or CBS_DROPDOWNLIST Or CBS_HASSTRINGS Or CBS_DISABLENOSCROLL, 10, 250, 140, 100, parent, IDC_COMBO, App.hInstance, 0)
    
    ' add some items to the listbox and combo box
    Dim i As Integer
    For i = 0 To 20
        SendMessage hWndList, LB_ADDSTRING, 0, "List Item " & i
        SendMessage hWndCombo, CB_ADDSTRING, 0, "List Item " & i
    Next i
    ' set a selection in the combo
    SendMessage hWndCombo, CB_SETCURSEL, 0, 0
    
    ' start a timer for fun, it will set off every second
    SetTimer parent, ID_TIMER, 1000, 0
End Sub

Private Function CreateNewWindow(ByVal MyWndProc As Long, ByVal szWindowClass As String, ByVal szWindowTitle As String, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long) As Long
    ' Register a class
    Dim wcex As WNDCLASSEX
    wcex.cbSize = LenB(wcex)
    wcex.style = CS_HREDRAW Or CS_VREDRAW Or CS_PARENTDC
    wcex.lpfnWndProc = MyWndProc
    wcex.cbClsExtra = 0
    wcex.cbWndExtra = 0
    wcex.hInstance = App.hInstance
    wcex.hIcon = 0
    wcex.hCursor = LoadCursor(0, IDC_ARROW)
    wcex.hbrBackground = COLOR_WINDOW + 1
    wcex.lpszMenuName = vbNullString
    wcex.lpszClassName = szWindowClass
    wcex.hIconSm = 0

    If RegisterClassEx(wcex) = 0 Then
        MsgBox "Failed to register window!"
        CreateNewWindow = -1
        Exit Function
    End If
    
    ' create the window
    Dim vbWindow As Long
    vbWindow = CreateWindowEx(WS_EX_APPWINDOW Or WS_EX_WINDOWEDGE, _
                              szWindowClass, _
                              szWindowTitle, _
                              WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPEDWINDOW, _
                              x, y, cx, cy, 0, 0, App.hInstance, 0)
                              
    If vbWindow = 0 Then
        MsgBox "Failed to create the window!"
        UnregisterClass szWindowClass, App.hInstance
        CreateNewWindow = -1
        Exit Function
    End If
    
    ' show the window
    UpdateWindow vbWindow
    ShowWindow vbWindow, SW_SHOW
    
    ' message loop to process window messages
    Dim myMsg As MSG
    While GetMessage(myMsg, 0, 0, 0) <> 0 ' waiting for PostQuitMessage to be called to break out
        TranslateMessage myMsg
        DispatchMessage myMsg
    Wend
    
    ' done with window.. clean up what we created
    DestroyWindow vbWindow
    UnregisterClass szWindowClass, App.hInstance
    
    ' return exit code
    CreateNewWindow = myMsg.wParam
End Function

Public Function HIWORD(dw As Long) As Integer
    If dw And &H80000000 Then
        HIWORD = (dw \ 65535) - 1
    Else
        HIWORD = dw \ 65535
    End If
End Function

Public Function LOWORD(dw As Long) As Integer
    If dw And &H8000& Then
        LOWORD = &H8000 Or (dw And &H7FFF&)
    Else
        LOWORD = dw And &HFFFF&
    End If
End Function

Public Function doWindow(ByVal szWindowTitle As String, ByVal szWindowClass As String, Optional ByVal x As Long = CW_USEDEFAULT, Optional ByVal y As Long = CW_USEDEFAULT, Optional ByVal cx As Long = CW_USEDEFAULT, Optional ByVal cy As Long = CW_USEDEFAULT) As Long
    doWindow = CreateNewWindow(AddressOf MainWndProc, szWindowClass, szWindowTitle, x, y, cx, cy)
End Function

