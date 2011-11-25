VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CLASS_NAME = "w>gopherlaunch"
Const GRIPPER_WIDTH = 16
Const BAR_HEIGHT = 30
Const ITEM_SIZE = 24
Const ITEM_GAP = 2
Const ICON_SIZE = 20

Dim mView As mfxView
Dim mhWnd As Long

Dim mConfig As String

Dim mExpanded As Boolean
Dim mItems As BTagList
Dim mCurrent As Long

Dim mColBackground As Long
Dim mTooltip As CToolTip
Dim mShading As Boolean

Dim mTrashAnim As TTrashWindow

Implements IDropTarget
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pwp As WINDOWPOS
Static fDragging As Boolean
Static fMenuOpen As Boolean
Static hWndPrevious As Long
Static rcwa As RECT
Static i As Long

'Static hDC As Long
'Static pps As PAINTSTRUCT

    Select Case uMsg

'    Case WM_ERASEBKGND
'        ReturnValue = -1
'        BWndProcSink_WndProc = True
'
'    Case WM_PAINT
'        hDC = BeginPaint(hWnd, pps)
'        draw_view mView, hDC
'        EndPaint hWnd, pps
'        ReturnValue = 0
'        BWndProcSink_WndProc = True


    Case WM_NCACTIVATE
        hWndPrevious = GetForegroundWindow()


'    Case WM_ACTIVATE
'        Debug.Print LoWord(wParam)
'        Select Case LoWord(wParam)
'        Case WA_ACTIVE, WA_CLICKACTIVE
'            Debug.Print "previous window was " & g_HexStr(lParam) & " / " & g_HexStr(GetFocus())
'
'        End Select
'
'    Case WM_MOUSEACTIVATE
'        ReturnValue = MA_NOACTIVATE
'        BWndProcSink_WndProc = True

    Case WM_NCHITTEST
        If g_IsPressed(vbKeyControl) Then
            ReturnValue = HTCAPTION
            BWndProcSink_WndProc = True
            Exit Function

        End If

    Case WM_LBUTTONDOWN
        uSetExpanded True
        SetCapture hWnd

    Case WM_LBUTTONUP
        i = uHitTest(new_BPointFromInt32(lParam))
        If i <> 0 Then
            If (wParam And MK_CONTROL) Then
                mItems.Remove i
                mCurrent = 0
                uArrange
                uRedraw
                uWriteConfig

                ' /* draw poof animation */

                With New TTrashWindow
                    .Go new_BPointFromInt32(GetMessagePos())

                End With

            Else
                uLaunch mItems.TagAt(i)

            End If
        End If

        ReleaseCapture

'    Case WM_LBUTTONDBLCLK
'        PostQuitMessage 0

    Case WM_CAPTURECHANGED
        uSetExpanded False
        If mCurrent <> 0 Then _
            mTooltip.Hide CStr(mCurrent)

        Debug.Print "previous window was " & g_HexStr(hWndPrevious)

        g_ShowWindow hWnd, False
        g_ShowWindow hWnd, True, False

'        If hWndPrevious Then _
'            g_WindowToFront hWndPrevious, True
'
'        hWndPrevious = 0


    Case WM_MOUSEMOVE
        i = uHitTest(new_BPointFromInt32(lParam))
        If i <> mCurrent Then
            If mCurrent <> 0 Then _
                mTooltip.Hide CStr(mCurrent)

            If i <> 0 Then _
                mTooltip.Show CStr(i)

            mCurrent = i
            uRedraw

        End If

    Case WM_RBUTTONDOWN
        If Not mExpanded Then
            uDoMenu

        Else
            ' /* possible expansion to allow for right-click context menu on the items? */

        End If

    Case WM_WINDOWPOSCHANGING
        If fDragging Then
            CopyMemory pwp, ByVal lParam, Len(pwp)
            If pwp.x < rcwa.Left Then
                pwp.x = rcwa.Left

            ElseIf (pwp.x + pwp.cx) > rcwa.Right Then
                pwp.x = rcwa.Right - pwp.cx

            End If

            If pwp.y < rcwa.Top Then
                pwp.y = rcwa.Top

            ElseIf (pwp.y + pwp.cy) > rcwa.Bottom Then
                pwp.y = rcwa.Bottom - pwp.cy

            End If

            CopyMemory ByVal lParam, pwp, Len(pwp)
            ReturnValue = 0
            BWndProcSink_WndProc = True

        End If

    Case WM_ENTERSIZEMOVE
        fDragging = True
        g_GetWorkArea rcwa

    Case WM_EXITSIZEMOVE
        fDragging = False
        uWriteConfig

    End Select

End Function

Private Sub Class_Initialize()

    EZRegisterClass CLASS_NAME

    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , WS_POPUP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    g_SetAlwaysOnTop mhWnd, True
    RegisterDragDrop mhWnd, Me

    Set mTooltip = New CToolTip
    mTooltip.Create TTS_ALWAYSTIP, mhWnd
'    mTooltip.Add "tip", TTF_TRACK Or TTF_TRANSPARENT

    Set mItems = new_BTagList()

    ' /* can specify custom configuration file */

    If Command$ <> "" Then
        mConfig = Command$

    Else
        mConfig = g_MakePath(App.Path) & "gopherlaunch.conf"

    End If

    mColBackground = rgba(244, 244, 244)

    ' /* read the config */

    uLoadConfig

'    uAdd "c:\windows\system32\notepad.exe"
'    uAdd "c:\windows\system32\cmd.exe"
'    uAdd "c:\windows\system32\calc.exe"

    ' /* create view, organise item frames and draw */

    Set mView = New mfxView
    uArrange
    uSetExpanded False

    ' /* make visible */

    g_ShowWindow mhWnd, True

End Sub

Private Sub Class_Terminate()

    mTooltip.Destroy
    Set mTooltip = Nothing

    RevokeDragDrop mhWnd
    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Sub uSetExpanded(ByVal IsExpanded As Boolean)

    mExpanded = IsExpanded
    uRedraw

End Sub

Private Sub uRedraw()
Dim pr As BRect

    With mView

        If mExpanded Then
            .SizeTo GRIPPER_WIDTH + (mItems.CountItems * (ITEM_SIZE + ITEM_GAP)) + 4, BAR_HEIGHT

        Else
            .SizeTo GRIPPER_WIDTH, BAR_HEIGHT

        End If

        .EnableSmoothing False
        .SetHighColour mColBackground
        .FillRect .Bounds

        ' /* gripper */
'        .SetHighColour rgba(130, 140, 191)
'
'        .FillRect new_BRect(0, 0, GRIPPER_WIDTH - 1, .Bounds.Bottom)


        ' /* gripper etching */
        
        Set pr = .Bounds
'        pr.Right = pr.Left + GRIPPER_WIDTH - 1
        pr.InsetBy 5, 5
        pr.Right = pr.Left

        .SetHighColour rgba(54, 54, 54, 88)
        .StrokeLine pr
        pr.OffsetBy 1, 0

        If mExpanded Then
            .SetHighColour rgba(255, 255, 255, 88)
            .StrokeLine pr

        End If

        pr.OffsetBy 3, 0
        .SetHighColour rgba(54, 54, 54, 88)
        .StrokeLine pr
        pr.OffsetBy 1, 0

        If mExpanded Then
            .SetHighColour rgba(255, 255, 255, 88)
            .StrokeLine pr

        End If


        ' /* shading */

        If mShading Then
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 48)
            .FillRect .Bounds, MFX_VERT_GRADIENT

        End If

Dim pb As mfxBitmap
Dim pi As TItem
Dim i As Long

        If mExpanded Then
            ' /* draw icons */
            If mItems.CountItems Then
                For i = 1 To mItems.CountItems
                    Set pi = mItems.TagAt(i)
                    Set pb = create_bitmap_from_image(pi.Image)

                    If i = mCurrent Then
                        .SetHighColour rgba(0, 0, 0, 32)
                        .FillRect pi.Frame.InsetByCopy(-1, -1)
                        .StrokeRect pi.Frame.InsetByCopy(-1, -1)
                        pb.ChangeBrightness -0.25

                    End If

                    .DrawScaledImage pb, pi.Frame.TopLeft.OffsetByCopy(Fix((ITEM_SIZE - ICON_SIZE) / 2), Fix((ITEM_SIZE - ICON_SIZE) / 2)), new_BPoint(ICON_SIZE, ICON_SIZE)

                Next i
            End If

            ' /* high/lowlight */

            .SetHighColour rgba(255, 255, 255, 56)
            .SetLowColour rgba(0, 0, 0, 56)
            .StrokeFancyRect .Bounds.InsetByCopy(1, 1)

'            ' /* splitter */
'            .SetHighColour rgba(0, 0, 0, 176)
'            .StrokeLine new_BRect(GRIPPER_WIDTH - 1, 0, GRIPPER_WIDTH - 1, .Bounds.Bottom)
'
'            ' /* high/lowlight */
'            .SetHighColour rgba(255, 255, 255, 56)
'            .SetLowColour rgba(0, 0, 0, 56)
'            .StrokeFancyRect new_BRect(0, 0, GRIPPER_WIDTH - 1, .Bounds.Bottom).InsetByCopy(1, 1)
'            .StrokeFancyRect new_BRect(GRIPPER_WIDTH - 1, 0, .Bounds.Right, .Bounds.Bottom).InsetByCopy(1, 1)

        End If


'        ' /* gripper etching */
'
'        Set pr = new_BRect(0, 4, 0, .Bounds.Bottom - 4)
'        pr.OffsetBy 5, 0
'
'        For i = 1 To 2
'            .StrokeLine pr, , MFX_SOLID_LOW
'            pr.OffsetBy 1, 0
'            .StrokeLine pr
'            pr.OffsetBy 3, 0
'
'        Next i

        ' /* border */
        .SetHighColour rgba(0, 0, 0, 196)
        .StrokeRect .Bounds

    End With

'    g_SizeWindow mhWnd, mView.Width, mView.Height
'    g_UpdateWindow mhWnd, False

    apply_view_to_window mView, mhWnd, IIf(mExpanded, 255, 180)

End Sub

Private Sub uArrange()

    If mItems.CountItems = 0 Then _
        Exit Sub

    mTooltip.RemoveAll

Static i As Long
Dim pr As BRect
Dim pi As TItem

    Set pr = new_BRect(0, 0, 23, 23)
    pr.OffsetBy GRIPPER_WIDTH + 2, ((BAR_HEIGHT - ITEM_SIZE) / 2)

    For i = 1 To mItems.CountItems
        Set pi = mItems.TagAt(i)
        Set pi.Frame = pr.Duplicate
        mTooltip.Add CStr(i), TTF_TRANSPARENT Or TTF_TRACK, pi.Label, pr.Duplicate
        pr.OffsetBy ITEM_SIZE + ITEM_GAP, 0

    Next i

End Sub

Private Function uHitTest(ByRef Point As BPoint) As Long

    If mItems.CountItems = 0 Then _
        Exit Function

Static i As Long
Dim pi As TItem

    For i = 1 To mItems.CountItems
        Set pi = mItems.TagAt(i)
        If pi.Frame.Contains(Point) Then
            uHitTest = i
            Exit Function

        End If

    Next i

End Function

'Private Function uGetIcon(ByRef Item As TItem) As MImage
'Dim pi As BIcon
'
'    With new_BIconContent(Item.Target)
'        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then _
'            Set uGetIcon = pi.Render()
'
'    End With
'
'End Function

Private Function uAdd(ByRef Info As CConfSection) As Boolean
Dim pi As TItem

    Set pi = New TItem
    If pi.SetFrom(Info) Then
        mItems.Add pi
        uAdd = True

    End If

End Function

Private Function uAddByTarget(ByVal Target As String) As Boolean
Dim pi As TItem

    Set pi = New TItem
    If pi.SetToTarget(Target) Then
        mItems.Add pi
        uAddByTarget = True

    End If

End Function

Private Sub uLaunch(ByRef Item As TItem)

    ShellExecute mhWnd, vbNullString, Item.Target, Item.Args, Item.WorkingDir, SW_SHOW

End Sub

Private Sub uWriteConfig()
Dim ps As CConfSection
Dim pi As TItem
Dim i As Long

    With New CConfFile3
        .SetFile mConfig

        Set ps = New CConfSection
        With ps
            .SetName "general"
            .Add "left", BW_Frame(mhWnd).Left
            .Add "top", BW_Frame(mhWnd).Top
            .Add "col-background", CStr(get_red(mColBackground)) & "," & CStr(get_green(mColBackground)) & "," & CStr(get_blue(mColBackground))
            .Add "shading", IIf(mShading, "1", "0")

        End With

        .Add ps

        If mItems.CountItems Then
            For i = 1 To mItems.CountItems
                Set ps = New CConfSection
                Set pi = mItems.TagAt(i)
                ps.SetName "item"
                pi.WriteConfig ps
                .Add ps

            Next i

        End If

        .Save

    End With

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)

End Sub

Private Sub IDropTarget_DragLeave()

End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)

End Sub

Private Sub IDropTarget_Drop(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pDrop As CDropContent
Dim pData As CDropItem
Dim sz As String

    Set pDrop = New CDropContent
    If pDrop.SetTo(pDataObject) Then
        With pDrop
            Debug.Print .CountFormats
            .Rewind
            Do While .GetNextFormat(sz)
                Debug.Print ": " & sz
            
            Loop

            If .HasFormat("FileNameW") Then
                If .GetData("FileNameW", pData) Then
                    Debug.Print "Added '" & pData.GetAsString(True) & "'"
                    uAddByTarget pData.GetAsString(True)
                    uWriteConfig
                    uArrange
                    uRedraw

                End If

            End If

        End With

    Else
        Debug.Print "IDropTarget.Drop(): failed to obtain IDataObject"

    End If

End Sub

Private Sub uLoadConfig()
Dim pt As POINTAPI

    ' /* defaults */

    pt.x = 0
    pt.y = 320
    mShading = True

Dim ps As CConfSection
Dim sz As String
Dim i As Long

    With New CConfFile3
        .SetFile mConfig
        If .Load() Then

            ' /* general section */

            If .Find("general", ps) Then
                pt.x = Val(ps.GetValueWithDefault("left", "0"))
                pt.y = Val(ps.GetValueWithDefault("top", "320"))
                i = g_RGBAFromString(ps.GetValueWithDefault("col-background"))
                If i <> 0 Then _
                    mColBackground = i

                If ps.Find("shading", sz) Then _
                    mShading = (sz = "1")

            End If

            ' /* item sections */

            .Rewind
            Do While .GetNextSection(ps)
                If ps.Name = "item" Then _
                    uAdd ps

            Loop

        End If

    End With

    ' /* position window */

    g_MoveWindow mhWnd, pt.x, pt.y

End Sub

Private Sub uDoMenu()
Dim pmi As OMMenuItem
Dim pm As OMMenu
Dim dw As Long

    With New OMMenu
        
        Set pm = New OMMenu
        With pm
            .AddItem .CreateItem("", "Drag items onto the gripper to add them; hold CTRL when invoking an item to delete it")
            .AddItem .CreateItem("", "Hold CTRL over the gripper to move the bar")
            .AddSeparator
            .AddItem .CreateItem("", "GopherLaunch " & App.Major & "." & App.Minor & " Build " & App.Revision & " " & App.LegalCopyright)

        End With

        .AddItem .CreateItem("", "About", , , , , , pm)
        .AddSeparator
        .AddItem .CreateItem("shad", "Shading", , , mShading)
        .AddItem .CreateItem("bcol", "Background colour...")
        .AddSeparator
        .AddItem .CreateItem("relc", "Reload config")
        .AddItem .CreateItem("erse", "Erase config")
        .AddSeparator
        .AddItem .CreateItem("quit", "Quit")

        Set pmi = .Track(mhWnd)
        If Not (pmi Is Nothing) Then
            Select Case pmi.Name
            Case "quit"
                PostQuitMessage 0

            Case "shad"
                mShading = Not mShading
                uWriteConfig
                uRedraw

            Case "erse"
                mItems.MakeEmpty
                uWriteConfig
                uLoadConfig
                uArrange
                uRedraw

            Case "relc"
                mItems.MakeEmpty
                mColBackground = rgba(244, 244, 244)
                uLoadConfig
                uArrange
                uRedraw

            Case "bcol"
                dw = g_MakeRGB24(mColBackground)
                If g_PickColour(dw, mhWnd) Then
                    mColBackground = g_MakeRGBA(dw)
                    uRedraw
                    uWriteConfig

                End If

'                mColBackground = rgba(Rnd * 255, Rnd * 255, Rnd * 255)

            End Select

        End If

    End With

End Sub

Private Function uItemAt(ByVal Index As Long) As TItem

    Set uItemAt = mItems.TagAt(Index)

End Function
