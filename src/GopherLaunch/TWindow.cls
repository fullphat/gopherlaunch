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

'Private Declare Function SetActiveWindow Lib "User32" (ByVal hwnd As Long) As Long

Const GRIPPER_WIDTH = 16
Const BAR_HEIGHT = 30
Const ITEM_SIZE = 24
Const ITEM_GAP = 2
Const ICON_SIZE = 20

Dim mView As mfxView
Dim mhWnd As Long

Dim mConfig As String
Dim mGeneral As CConfSection

Dim mExpanded As Boolean
Dim mItems As BTagList
Dim mCurrent As Long
'Dim mColBackground As Long
Dim mTooltip As CToolTip
'Dim mShading As Boolean
Dim mTrashAnim As TTrashWindow
Dim thePanel As BPrefsPanel

Implements BWndProcSink
Implements KPrefsPanel
Implements KPrefsPage
Implements IDropTarget

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print Control.GetName

'    Select Case Control.GetName
'    Case "type"
'        prefskit_SafeEnable thePanel, "col2", (Value = "2")
'        prefskit_SafeEnable thePanel, "image-file", (Value = "3")
'        prefskit_SafeEnable thePanel, "use_random_pic", (Value = "3")
'        prefskit_SafeEnable thePanel, "image-pos", (Value = "3")
'        prefskit_SafeEnable thePanel, "change_pic", (mGeneral.GetValueWithDefault("type") = "3") And (mGeneral.GetValueWithDefault("use_random_pic") = "1")
'
'    Case "use_random_pic"
'        prefskit_SafeEnable thePanel, "change_pic", (mGeneral.GetValueWithDefault("type") = "3") And (mGeneral.GetValueWithDefault("use_random_pic") = "1")
'        If Value = "1" Then _
'            mImageToUse = uGetRandomFile(mGeneral.GetValueWithDefault("image-file"))
'
'    Case "go_site", "change_pic"
'        Exit Sub
'
'    End Select

    mGeneral.Update Control.GetName, Value
    uWriteConfig
    uRedraw

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "go_site"
        ShellExecute mhWnd, "open", "http://www.fullphat.net", vbNullString, vbNullString, SW_SHOW

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hwnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set thePanel = Nothing
    uSetExpanded False
    EnableWindow mhWnd, -1

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pwp As WINDOWPOS
Static fDragging As Boolean
Static fMenuOpen As Boolean
Static rcwa As RECT
Static i As Long

'Static hDC As Long
'Static pps As PAINTSTRUCT

    Select Case uMsg

    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_SETCURSOR
        If (LoWord(lParam) = -2) And (HiWord(lParam) = WM_LBUTTONDOWN) And (NOTNULL(thePanel)) Then
            g_WindowToFront thePanel.hWnd, True
            ReturnValue = -1
            BWndProcSink_WndProc = True

        End If

'    Case WM_ERASEBKGND
'        ReturnValue = -1
'        BWndProcSink_WndProc = True
'
'
'    Case WM_PAINT
'        hDC = BeginPaint(hWnd, pps)
'        draw_view mView, hDC
'        EndPaint hWnd, pps
'        ReturnValue = 0
'        BWndProcSink_WndProc = True


    Case WM_NCHITTEST
        If g_IsPressed(vbKeyControl) Then
            ReturnValue = HTCAPTION
            BWndProcSink_WndProc = True

        End If


    Case WM_RBUTTONDOWN
        If Not mExpanded Then
            SetForegroundWindow hWnd
            uDoMenu

        Else
            ' /* possible expansion to allow for right-click context menu on the items? */

        End If


    Case WM_LBUTTONDOWN
        uSetExpanded True
        SetCapture hWnd
        ReturnValue = 0
        BWndProcSink_WndProc = True


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


    Case WM_CAPTURECHANGED
        uSetExpanded False
        If mCurrent <> 0 Then _
            mTooltip.Hide CStr(mCurrent)

'        Debug.Print "previous window was " & g_HexStr(hWndPrevious)


    Case WM_SYSCOMMAND
        Select Case wParam
        Case SC_KEYMENU
            uDoMenu new_BPoint(BW_Frame(hWnd).Left, BW_Frame(hWnd).Bottom)

        End Select


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
        mGeneral.Update "left", BW_Frame(mhWnd).Left
        mGeneral.Update "top", BW_Frame(mhWnd).Top
        uWriteConfig

    End Select

End Function

Private Sub Class_Initialize()

    EZRegisterClass CLASS_NAME

    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , WS_POPUP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, WS_EX_TOOLWINDOW Or WS_EX_TOPMOST Or WS_EX_NOACTIVATE)
    RegisterDragDrop mhWnd, Me

    Set mTooltip = New CToolTip
    mTooltip.Create TTS_ALWAYSTIP, mhWnd
    mTooltip.Add "tip", TTF_TRACK Or TTF_TRANSPARENT

    Set mItems = new_BTagList()

    ' /* can specify custom configuration file */

    If Command$ <> "" Then
        mConfig = Command$

    Else
        mConfig = g_MakePath(App.Path) & "gopherlaunch.conf"

    End If

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
        .SetHighColour g_SafeLong(mGeneral.GetValueWithDefault("col-background"))
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

        If mGeneral.GetValueWithDefault("shading") = "1" Then
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

        If (mExpanded) And (mGeneral.GetValueWithDefault("glass") = "1") Then
            ' /* glass */
            .SetHighColour rgba(255, 255, 255, 30)
            .SetLowColour rgba(255, 255, 255, 80)
            .FillRect new_BRect(0, 0, .Bounds.Right, Fix(.Bounds.Bottom / 2)), MFX_VERT_GRADIENT

        End If

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

        .Add mGeneral

'        Set ps = New CConfSection
'        With ps
'            .SetName "general"
'            .Add "left", BW_Frame(mhWnd).Left
'            .Add "top", BW_Frame(mhWnd).Top
'            .Add "col-background", CStr(get_red(mColBackground)) & "," & CStr(get_green(mColBackground)) & "," & CStr(get_blue(mColBackground))
'            .Add "shading", IIf(mShading, "1", "0")
'
'        End With
'
'        .Add ps

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
Dim ps As CConfSection
Dim sz As String
'Dim i As Long

    With New CConfFile3
        .SetFile mConfig
        If .Load() Then

            ' /* general section */

            .Find "general", mGeneral

            ' /* item sections */

            .Rewind
            Do While .GetNextSection(ps)
                If ps.Name = "item" Then _
                    uAdd ps

            Loop

        End If

    End With

    If ISNULL(mGeneral) Then
        Set mGeneral = New CConfSection
        mGeneral.SetName "general"

    End If

    With mGeneral

        ' /* validate [general] settings */

        .AddIfMissing "left", "0"
        .AddIfMissing "top", "320"
        .AddIfMissing "col-background", CStr(rgba(244, 244, 244))
        .AddIfMissing "shading", "1"
        .AddIfMissing "glass", "0"

        ' /* position window */

        g_MoveWindow mhWnd, g_SafeLong(.GetValueWithDefault("left")), g_SafeLong(.GetValueWithDefault("top"))

    End With

End Sub

Private Sub uDoMenu(Optional ByRef Position As BPoint)
Dim pmi As OMMenuItem
'Dim pm As OMMenu
Dim dw As Long

    With New OMMenu
        
'        Set pm = New OMMenu
'        With pm
'            .AddItem .CreateItem("", "Drag items onto the gripper to add them; hold CTRL when invoking an item to delete it")
'            .AddItem .CreateItem("", "Hold CTRL over the gripper to move the bar")
'            .AddSeparator
'            .AddItem .CreateItem("", "GopherLaunch " & App.Major & "." & App.Minor & " Build " & App.Revision & " " & App.LegalCopyright)
'
'        End With

'        .AddItem .CreateItem("", "About", , , , , , pm)
'        .AddSeparator
        .AddItem .CreateItem("pref", "Settings...")
'        .AddItem .CreateItem("shad", "Shading", , , (mGeneral.GetValueWithDefault("shading") = "1"))
'        .AddItem .CreateItem("bcol", "Background colour...")
'        .AddSeparator
'        .AddItem .CreateItem("relc", "Reload config")
'        .AddItem .CreateItem("erse", "Erase config")
        .AddSeparator
        .AddItem .CreateItem("quit", "Close")

        Set pmi = .Track(mhWnd, Position)
        If Not (pmi Is Nothing) Then
            Select Case pmi.Name
            Case "quit"
                PostQuitMessage 0

            Case "pref"
                uDoPrefs

            Case "shad"
                mGeneral.Update "shading", IIf(mGeneral.GetValueWithDefault("shading") = "1", "0", "1")
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
'                mColBackground = rgba(244, 244, 244)
                uLoadConfig
                uArrange
                uRedraw

            Case "bcol"
'                dw = g_MakeRGB24(mColBackground)
'                If g_PickColour(dw, mhWnd) Then
'                    mColBackground = g_MakeRGBA(dw)
'                    uRedraw
'                    uWriteConfig
'
'                End If
'
''                mColBackground = rgba(Rnd * 255, Rnd * 255, Rnd * 255)

            End Select

        End If

    End With

End Sub

Private Function uItemAt(ByVal Index As Long) As TItem

    Set uItemAt = mItems.TagAt(Index)

End Function

Public Sub uDoPrefs()
Dim pp As BPrefsPage
Dim pm As CTempMsg

    If ISNULL(thePanel) Then

        g_Debug "uDoPrefs(): creating panel..."

        Set thePanel = New BPrefsPanel
        With thePanel
            .SetHandler Me

            thePanel.SetTitle "GopherLaunch Preferences"
            thePanel.SetWidth 400

            ' /* general page */

            g_Debug "frmAbout.NewDoPrefs(): general page..."
            Set pp = new_BPrefsPage("Settings", load_image_obj(g_MakePath(App.Path) & "ico\prefs-general.png"), Me)
            With pp
                .SetMargin 96

                ' /* launch options */

                .Add new_BPrefsControl("fancycolourpicker", "col-background", , "Background:", mGeneral.GetValueWithDefault("col-background"))
                .Add new_BPrefsControl("fancytoggle2", "shading", "Enable shading?", , mGeneral.GetValueWithDefault("shading"))
                .Add new_BPrefsControl("fancytoggle2", "glass", "Enable glass effect?", , mGeneral.GetValueWithDefault("glass"))

'                .Add new_BPrefsControl("banner", "", "Background")
'                .Add new_BPrefsControl("fancycycle", "type", "Solid|Gradient|Image|Desktop", "Type:", mGeneral.GetValueWithDefault("type"))
'                .Add new_BPrefsControl("fancycolourpicker", "col2", , "Secondary colour:", mGeneral.GetValueWithDefault("col2"), , (mGeneral.GetValueWithDefault("type") = "2"))
'                .Add new_BPrefsControl("fancyfilepicker", "image-file", mGeneral.GetValueWithDefault("image-file"), "Image:", , , (mGeneral.GetValueWithDefault("type") = "3"))
'                .Add new_BPrefsControl("fancycycle", "image-pos", "Centre|Tile|Scale|Fit", "Position:", mGeneral.GetValueWithDefault("image-pos"), , (mGeneral.GetValueWithDefault("type") = "3"))
'                .Add new_BPrefsControl("banner", "", "Options")
'                .Add new_BPrefsControl("fancybutton2", "change_pic", "Change", "", , , (mGeneral.GetValueWithDefault("type") = "3") And (mGeneral.GetValueWithDefault("use_random_pic") = "1"))

            End With
            .AddPage pp

'            Set pp = new_BPrefsPage("Items", load_image_obj(g_MakePath(App.Path) & "ico\prefs-mark.png"), Me)
'            With pp
'                .SetMargin 128
''                .Add new_BPrefsControl("banner", "", "Watermark")
''                .Add new_BPrefsControl("fancytoggle2", "use_watermark", "Draw watermark?", , mGeneral.GetValueWithDefault("use_watermark"))
''                .Add new_BPrefsControl("fancyfilepicker", "watermark-file", mGeneral.GetValueWithDefault("watermark-file"), "Image:")
''                .Add new_BPrefsControl("fancycycle", "watermark-pos", "Top left|Top right|Bottom left|Bottom right", "Position:", mGeneral.GetValueWithDefault("watermark-pos"))
'
'            End With
'            .AddPage pp


'            Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "ico\prefs-extr.png"), Me)
'            With pp
'                .SetMargin 24
''                .Add new_BPrefsControl("banner", "", "Other Settings")
''                .Add new_BPrefsControl("fancytoggle2", "obscure_taskbar", "Obscure Taskbar?", , mGeneral.GetValueWithDefault("obscure_taskbar"), , False)
'
'            End With
'            .AddPage pp


            Set pp = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "ico\prefs-about.png"), Me)
            With pp
                .SetMargin 24
                .Add new_BPrefsControl("banner", "", "About GopherLaunch")
                Set pm = New CTempMsg
                pm.Add "image-file", g_MakePath(App.Path) & "app.png"
                pm.Add "image-height", 48
                pm.Add "valign", "centre"
                .Add new_BPrefsControl("labelex", "", "R" & App.Major & "." & App.Minor & " " & App.Comments & " (Build " & App.Revision & ")", , , pm)
                .Add new_BPrefsControl("label", "", "© 2011 full phat products")
                .Add new_BPrefsControl("label", "", "http://www.fullphat.net/")
                .Add new_BPrefsControl("label", "", "Released under the Simplified BSD Licence")
                .Add new_BPrefsControl("fancybutton2", "go_site", "Web site")
                
                .Add new_BPrefsControl("banner", "", "Hints & Tips")
                .Add new_BPrefsControl("label", "", "•  Drag items onto the gripper to add them")
                .Add new_BPrefsControl("label", "", "•  Hold CTRL when invoking an item to delete it")
                .Add new_BPrefsControl("label", "", "•  Hold CTRL over the gripper to move the bar")

            End With
            .AddPage pp


'            ' /* apps */
'
'            g_Debug "frmAbout.NewDoPrefs(): apps page..."
'            Set mAppsPage = New TAppsPage
'            .AddPage new_BPrefsPage("Apps", load_image_obj(g_MakePath(App.Path) & "etc\icons\apps.png"), mAppsPage)

'            Set pp = new_BPrefsPage("Styles", load_image_obj(g_MakePath(App.Path) & "etc\icons\styles.png"), Me)
'            With pp
'                .SetMargin 0
'                Set pm = New CTempMsg
'                pm.Add "height", 412
'                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
'                BTabStrip_AddPage pc, "Display Styles", new_BPrefsPage("sty-display", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Redirect Styles", new_BPrefsPage("sty-redirect", , New TNetSubPage)
'                .Add pc
'
'            End With
'            .AddPage pp
'
'                .Add new_BPrefsControl("fancytoolbar", "ftb>web_stuff", "Snarl Website|Discussion Group|Blog")
'
'
            g_Debug "frmAbout.NewDoPrefs(): displaying..."
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd

            uSetExpanded True
            EnableWindow mhWnd, 0

            g_Debug "frmAbout.NewDoPrefs(): done"

        End With

    End If

    g_ShowWindow thePanel.hWnd, True, True
    SetForegroundWindow thePanel.hWnd

End Sub

