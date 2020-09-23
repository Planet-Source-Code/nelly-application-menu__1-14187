VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApplication 
   Caption         =   "Application Menu. (c) Neil Etherington 2000."
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   105
      TabIndex        =   5
      Top             =   6360
      Width           =   8535
      Begin VB.TextBox txtInfo 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton cmdAddInfo 
         Caption         =   "Add Application:"
         Height          =   530
         Left            =   120
         Picture         =   "frmApplication.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel:"
         Height          =   530
         Left            =   1800
         Picture         =   "frmApplication.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         Caption         =   "Main"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   345
      End
   End
   Begin MSComctlLib.ListView lstPath 
      Height          =   6015
      Left            =   4785
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10610
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Do not edit any of these entries."
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   225
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplication.frx":0294
            Key             =   "appcldfld"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplication.frx":03EE
            Key             =   "appopnfld"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplication.frx":0548
            Key             =   "appcldfle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmApplication.frx":06A2
            Key             =   "appopnfle"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView AppTree 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu h1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuApplication 
      Caption         =   "Application"
      Begin VB.Menu h2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuApplicationAdd 
         Caption         =   "Application Add:"
      End
      Begin VB.Menu mnuApplicationRemove 
         Caption         =   "Application Remove:"
      End
   End
   Begin VB.Menu mnuCategory 
      Caption         =   "Category"
      Begin VB.Menu h3 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCategoryAdd 
         Caption         =   "Category Add:"
      End
      Begin VB.Menu mnuCategoryRemove 
         Caption         =   "Category Remove:"
      End
   End
End
Attribute VB_Name = "frmApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SW_SHOWNORMAL = 1

'***********************************************************************************
'Form Events.
'***********************************************************************************
'Form Form_Load.
Private Sub Form_Load()
    
    Dim cApplication As cApplication
        Set cApplication = New cApplication
    
    'Open Application Entries.------------------------------------------------------
    cApplication.subOpenApplicationEntries
    '-------------------------------------------------------------------------------

End Sub

'Form Form_Activate.
Private Sub Form_Activate()
        
    'Disable Menu items.------------------------------------------------------------
    With frmApplication
        .mnuApplicationAdd.Enabled = False
        .mnuApplicationRemove.Enabled = False
        .mnuCategoryRemove.Enabled = False
    End With
    '-------------------------------------------------------------------------------
        
End Sub
 
'Form Form_Queryunload.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim cApplication As cApplication
        Set cApplication = New cApplication
    
    'Save Application Entries.------------------------------------------------------
    cApplication.subSaveApplicationEntries
    '-------------------------------------------------------------------------------

End Sub

'Form Form_Unload.
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frmApplication = Nothing
End Sub


'***********************************************************************************
'AppTree Events.
'***********************************************************************************
'AppTree AppTree_Click.
Private Sub AppTree_Click()
'
End Sub

'AppTree AppTree_KeyDown.
Private Sub AppTree_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'lstPath is Invisible by default but this combination of Keys makes it visible--
    'This is for Test purposes.
    Dim Ctrl
    Ctrl = (Shift And vbCtrlMask)
    
    If Ctrl And KeyCode = vbKeyF12 Then
        With lstPath
            If .Visible = False Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
    End If

End Sub

'AppTree AppTree_NodeClick.
Private Sub AppTree_NodeClick(ByVal Node As MSComctlLib.Node)

    AppTree.Nodes(AppTree.SelectedItem.Index).Selected = True
    
    'Enable and Disable menu items(Determined by what item in AppTree is selected).-
    With AppTree
        If .Nodes.Count > 0 Then
            If .SelectedItem.Text = .SelectedItem.FullPath Then
                With frmApplication
                    .mnuApplicationAdd.Enabled = True
                    .mnuApplicationRemove.Enabled = False
                    .mnuCategoryRemove.Enabled = False
                    If AppTree.SelectedItem.Children = 0 Then
                        .mnuCategoryRemove.Enabled = True
                    End If
                End With
            ElseIf .SelectedItem.Text <> .SelectedItem.FullPath Then
                With frmApplication
                    .mnuApplicationAdd.Enabled = False
                    .mnuApplicationRemove.Enabled = True
                    .mnuCategoryRemove.Enabled = False
                End With
            End If
        End If
    End With
    '-------------------------------------------------------------------------------
    
End Sub

'AppTree AppTree_DblClick.
Private Sub AppTree_DblClick()
    
    On Error GoTo errHandler
    
    Dim sReturn As String
    Dim sEntry As String
    Dim XFDll As Object
        Set XFDll = CreateObject("cXFileP.cXFileFM")
    Dim sPath As String
    Dim lShellApp As Long
    Dim sFile As String
    Dim sCPath As String
    
    'If the User selects anything but a "root" item, then search lstPath for the----
    'selected item Path. Once found Shell Application.
    With AppTree
        If .SelectedItem.Text <> .SelectedItem.FullPath Then
            sEntry$ = lstPath.FindItem(.SelectedItem.Text, , , lvwPartial)
        
            If sEntry$ > Trim("") Then
                Me.MousePointer = 11
                'To Shell an Application we need to know the Filename and correct Path
                'of the selected File, to do this we first retrieve the Path and Name
                'of the File in lstPath eg.("c:\WinNT\explorer.exe"). We then extract
                'the Path and Name from sPath$, so sFile$ becomes ("explorer.exe") and
                'sCPath$ becomes ("c:\WinNT\"). We can now Shell the Application.
                sPath$ = XFDll.ftnShellApplication(sEntry$ & "\")
                sFile$ = XFDll.ftnFolderRename(sPath$, "Last")
                sCPath$ = XFDll.ftnFolderRename(sPath$, "First")
                lShellApp& = ShellExecute(Me.hwnd, "open", sFile$, vbNullString, sCPath$, SW_SHOWNORMAL)
                Me.MousePointer = 0
                Form_Unload (0)
                '-------------------------------------------------------------------
            End If
                        
        End If
    End With
    Exit Sub
    '-------------------------------------------------------------------------------
    
errHandler:
    Select Case Err.Number
        'Dbl Check, If the Path of the selected item is not found, then notify user-
        'and ask if they want the selected item removing from AppTree and lstPath.
        Case 53
            Me.MousePointer = 0
            sReturn$ = MsgBox("Run time error '53'" & vbCrLf & vbCrLf _
            & "File not found" & vbCrLf & vbCrLf _
            & "Remove this item..", vbCritical + vbYesNo, App.EXEName)
            If sReturn$ = vbYes Then
                Call mnuApplicationRemove_Click
            End If
            Exit Sub
        '---------------------------------------------------------------------------
        
        Case Else
            Me.MousePointer = 0
            MsgBox "frmApplication. AppTree_SbdClick." & vbCrLf & vbCrLf _
            & "Err.Number. " & Err.Number & vbCrLf _
            & "Err.Description. " & Err.Description, vbCritical + vbOKOnly, "X-File:"
            Exit Sub
    End Select

End Sub


'***********************************************************************************
'Command cmdAddInfo Events.
'***********************************************************************************
'Command Command_Click.
Private Sub cmdAddInfo_Click()

    On Error GoTo errHandler
    
    Dim nNode As Node
    Dim lstItem As ListItem
    Dim XFDll As Object
        Set XFDll = CreateObject("cXFileP.cXFileFM")
    Dim sFileExists As String

    'If nothing is entered into txtinfo or txtPath then Exit Sub.-------------------
    If cmdAddInfo.Caption = "Add Category." Then
        If txtInfo.Text = Trim("") Then
            MsgBox "Nothing to add.", vbCritical + vbOKOnly, "X-File:"
        Exit Sub
        End If
    End If
    
    If cmdAddInfo.Caption = "Add Application." Then
        If txtInfo.Text = Trim("") Or txtPath.Text = Trim("") Then
            MsgBox "Nothing to add.", vbCritical + vbOKOnly, "X-File:"
        Exit Sub
    End If
    End If
    '-------------------------------------------------------------------------------
    
    'Check to see if a txtInfo.Text(Name of Application) is already resident in AppTree
    For Each nNode In AppTree.Nodes
        If nNode.Text = txtInfo.Text Then
            MsgBox "An entry already has the Name of... " & "'" & txtInfo.Text & "'", vbCritical + vbOKOnly, "X-File:"
        Exit Sub
        End If
    Next nNode
    '-------------------------------------------------------------------------------

    'Check to see if File does exist before adding it to AppTree.-------------------
    If cmdAddInfo.Caption = "Add Application." Then
        sFileExists$ = XFDll.ftnFileExists(txtPath.Text)
        If sFileExists$ = "False" Then
            MsgBox "The File " & "'" & txtInfo.Text & "'" & " cannot be found at " & "'" & txtPath.Text & "'" & vbCrLf & vbCrLf _
            & "Please check your Path..", vbCritical + vbOKOnly, "X-File:"
        Exit Sub
        End If
    End If

    With AppTree
        If cmdAddInfo.Caption = "Add Category." Then
            Set nNode = .Nodes.Add(, , "root" & txtInfo.Text, txtInfo.Text, "appcldfld", "appopnfld")
        ElseIf cmdAddInfo.Caption = "Add Application." Then
            .Nodes.Add .SelectedItem.Key, tvwChild, .SelectedItem.Key & txtInfo.Text, txtInfo.Text, "appcldfle", "appopnfle"
            Set lstItem = lstPath.ListItems.Add(, , txtInfo.Text & "(~" & txtPath.Text & "~)")
        End If
    End With
    Exit Sub
    '-------------------------------------------------------------------------------

errHandler:
    Select Case Err.Number
        Case 35602
            MsgBox "Run time error 35602" & vbCrLf & vbCrLf _
            & "Key is not unique in collection", vbCritical + vbOKOnly, "X-File:"
            Exit Sub
        Case 91
            MsgBox "Please select a Category before adding an Application.", vbCritical + vbOKOnly, "X-File:"
            Exit Sub
        Case Else
            MsgBox "frmApplication. cmdAddIfo." & vbCrLf & vbCrLf _
            & "Err.Number" & Err.Number & vbCrLf _
            & "Err.Description" & Err.Description, vbCritical + vbOKOnly, "X-File:"
            Exit Sub
    End Select
End Sub

'Cancel.
Private Sub cmdCancel_Click()
    Form_Unload (0)
End Sub


'***********************************************************************************
'Menu File Commands.
'***********************************************************************************
'Menu mnuFileExit_Click.
Private Sub mnuFileExit_Click()
    Form_Unload (0)
End Sub


'***********************************************************************************
'Menu Category Commands.
'***********************************************************************************
'Menu mnuCategoryAdd_Click.
Private Sub mnuCategoryAdd_Click()
    Me.Height = 8835
    txtPath.Text = ""
    txtPath.Visible = False
    txtInfo.Width = 8295
    lblMain.Caption = "Please enter Category Name, (No Duplicate Names are allowed)."
    cmdAddInfo.Caption = "Add Category."
    txtInfo.Text = ""
    txtInfo.SetFocus
End Sub

'Menu mnuCateGoryRemove_Click.
Private Sub mnuCategoryRemove_Click()
'    On Error Resume Next
    'Remove Category from AppTree.--------------------------------------------------
    AppTree.Nodes.Remove (AppTree.SelectedItem.Index)
    mnuCategoryRemove.Enabled = False
    '-------------------------------------------------------------------------------
End Sub

'***********************************************************************************
'Menu Application Commands.
'***********************************************************************************
'Menu mnuApplicationAdd_Click.
Private Sub mnuApplicationAdd_Click()
    Me.Height = 8835
    txtPath.Visible = True
    txtPath.Text = "Please enter full Path and Name of Application here."
    txtInfo.Width = 4095
    lblMain.Caption = "Please enter Application Name, (No Duplicate Names are allowed)."
    cmdAddInfo.Caption = "Add Application."
    txtInfo.Text = "Please enter description of Application here."
    txtInfo.SelLength = Len(txtInfo.Text)
    txtInfo.SetFocus
End Sub

'Menu mnuApplicationRemove_Click.
Private Sub mnuApplicationRemove_Click()
    
    'Remove Application from AppTree.-----------------------------------------------
    With AppTree
        If .SelectedItem.Text <> .SelectedItem.FullPath And .SelectedItem.Children = 0 Then
            Dim lstEntry As ListItem
                Set lstEntry = lstPath.FindItem(.SelectedItem.Text, , , lvwPartial)
            lstPath.ListItems.Remove (lstEntry.Index)
            .Nodes.Remove (.SelectedItem.Index)
        ElseIf .SelectedItem.Text = .SelectedItem.FullPath And .SelectedItem.Children = 0 Then
            .Nodes.Remove (.SelectedItem.Index)
        End If
    End With
    '-------------------------------------------------------------------------------

End Sub







