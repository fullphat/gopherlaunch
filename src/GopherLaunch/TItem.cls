VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Frame As BRect

Dim mTarget As String
Dim mIcon As String
Dim mFolder As String
Dim mArgs As String

Dim mIconGfx As MImage
Dim mLabel As String

Implements BTagItem

Private Function BTagItem_Name() As String
End Function

Private Function BTagItem_Value() As String
End Function

Public Function SetFrom(ByRef Info As CConfSection) As Boolean

    mTarget = Info.GetValueWithDefault("target")

    If mTarget = "" Then _
        Exit Function

Dim psl As CShellLink

    Select Case g_GetExtension(mTarget, True)
    Case "lnk"
        Set psl = New gopherlaunch.CShellLink
        If psl.SetTo(mTarget) Then _
            mTarget = psl.Target

    End Select

    ' /* optional */

    mIcon = Info.GetValueWithDefault("icon")
    mFolder = Info.GetValueWithDefault("dir")
    mArgs = Info.GetValueWithDefault("args")
    mLabel = Info.GetValueWithDefault("label")

    If mLabel = "" Then _
        mLabel = g_RemoveExtension(g_FilenameFromPath(mTarget))

    ' /* icon */

Dim pc As BIconContent

    If (mIcon = "") And (g_IsFolder(mTarget)) Then
        Set mIconGfx = load_image_obj(g_MakePath(App.Path) & "default-folder.png")

    Else
        Set pc = new_BIconContent(IIf(mIcon <> "", mIcon, mTarget))
        Set mIconGfx = uGetIcon(pc)

    End If

    If (mIconGfx Is Nothing) Then _
        Set mIconGfx = load_image_obj(g_MakePath(App.Path) & "app.png")

    SetFrom = True

End Function

Public Function SetToTarget(ByVal Target As String) As Boolean

    If Target = "" Then _
        Exit Function

Dim psl As CShellLink
Dim sz As String

    mTarget = Target

    Select Case g_GetExtension(Target, True)
    Case "lnk"

        ' /* peek inside the shortcut for the actual target */

        Set psl = New gopherlaunch.CShellLink
        If psl.SetTo(Target) Then

            If psl.CustomIcon <> "" Then _
                mIcon = psl.CustomIcon

            If psl.WorkingDir <> "" Then _
                mFolder = psl.WorkingDir

            If psl.Args <> "" Then _
                mArgs = psl.Args

'            If psl.Description <> "" Then
'                mLabel = psl.Description
'
'            Else
                mLabel = g_RemoveExtension(g_FilenameFromPath(Target))
'
'            End If

            ' /* get the real target */

            mTarget = psl.Target

        End If

    Case Else

        ' /* set label from filename */

        mLabel = g_RemoveExtension(g_FilenameFromPath(Target))

    End Select

    ' /* figure out the icon */

Dim pc As BIconContent
Dim pi As BIcon

    If (mIcon = "") And (g_IsFolder(mTarget)) Then
        Set mIconGfx = load_image_obj(g_MakePath(App.Path) & "default-folder.png")

    Else

        If (mIcon <> "") Or (g_GetExtension(mTarget, True) = "exe") Then
            ' /* extract icon from executable or custom icon field */
            sz = LCase$(IIf(mIcon <> "", mIcon, mTarget))
            If g_SafeRightStr(sz, 6) = ".ico,0" Then
                ' /* load from the ico */
                Set pc = New BIconContent
                If pc.LoadFromICO(g_SafeLeftStr(sz, Len(sz) - 2)) Then _
                    Set mIconGfx = uGetIcon(pc)

            Else
                ' /* load from the resource */
                Set pc = new_BIconContent(sz)
                Set mIconGfx = uGetIcon(pc)

            End If

            If (mIconGfx Is Nothing) Then _
                Set mIconGfx = load_image_obj(g_MakePath(App.Path) & "default-app.png")

        Else
            ' /* set icon from mime type */
            sz = g_MimeTypeFromExtension2(g_GetExtension(mTarget))
            Debug.Print "MIME: " & sz

        End If

        If (mIconGfx Is Nothing) Then _
            Set mIconGfx = load_image_obj(g_MakePath(App.Path) & "default-file.png")

    End If

    SetToTarget = True

End Function

Public Function Target() As String

    Target = mTarget

End Function

Public Function Icon() As String

    Icon = mIcon

End Function

Public Function WorkingDir() As String

    WorkingDir = mFolder

End Function

Public Function Args() As String

    Args = mArgs

End Function

Public Sub WriteConfig(ByRef Info As CConfSection)

    With Info
        .Add "target", mTarget

        If mIcon <> "" Then _
            .Add "icon", mIcon

        If mFolder <> "" Then _
            .Add "dir", mFolder
    
        If mArgs <> "" Then _
            .Add "args", mArgs

        If mLabel <> "" Then _
            .Add "label", mLabel

    End With

End Sub

Public Function Image() As MImage

    Set Image = mIconGfx

End Function

Public Function Label() As String

    Label = mLabel

End Function

Private Function uGetIcon(ByRef Content As BIconContent) As MImage
Dim pc As BIconContent
Dim pi As BIcon

    ' /* want the 32x32 icon ideally */

    Set pc = Content.Filter(B_PROPERTY_WIDTH, 32, B_SCOPE_EXACT)

    Debug.Print Content.CountImages & " > " & pc.CountImages

    If pc.CountImages = 1 Then
        ' /* if there's only one, we'll use that */
        Set uGetIcon = pc.IconAt(1).Render
        Exit Function

    ElseIf pc.CountImages > 1 Then
        ' /* otherwise, use the one with the greater colour depth */
        If Not pc.Find(B_PROPERTY_COLOUR_DEPTH, 32, B_SCOPE_EXACT, pi) Then _
            pc.Find B_PROPERTY_COLOUR_DEPTH, 24, B_SCOPE_LESS_THAN_OR_EQUAL_TO, pi

        If Not (pi Is Nothing) Then
            Debug.Print "best colour depth is " & pi.ColourDepth
            Set uGetIcon = pi.Render
            Exit Function

        End If

    End If

    ' /* if still nothing, use this more general approach */

    If pc.GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then
        Set uGetIcon = pi.Render()

    ElseIf Content.CountImages = 1 Then
        Set uGetIcon = Content.IconAt(1).Render

    End If

End Function
