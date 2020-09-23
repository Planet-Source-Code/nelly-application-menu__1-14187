VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'***********************************************************************************
'Open Application Menu Entries
'***********************************************************************************
Public Sub subOpenApplicationEntries()

    On Error Resume Next

    Dim sDel As String
    Dim sAppTree As String
    Dim vAppTree As Variant
    Dim iFree As Integer
        iFree = FreeFile
    Dim lstItem As ListItem
    Dim sAppPath As String
        sAppPath$ = mProcFunc.ftnAppPath(App.Path)
        
    sDel$ = "!"

    'open the Application Entries and add them to frmApplication.AppTree.-----------
    Open sAppPath$ & "Application.txt" For Input As #iFree
    
    Do
        
        Line Input #iFree, sAppTree$
        
            'Use Split Command to extract AppTree items.----------------------------
            vAppTree = Split(sAppTree$, sDel$)
            '-----------------------------------------------------------------------
    
        'If we have a "root" Item, then Add to AppTree using vAppTree(Array) information
        If vAppTree(2) = "root" Then
            Set nNode = frmApplication.AppTree.Nodes.Add(, , vAppTree(1), vAppTree(0), vAppTree(3), vAppTree(4))
            nNode.Expanded = True
        '---------------------------------------------------------------------------
        
        'If its a Sub Item then Add to AppTree as you would normally.---------------
        Else
            frmApplication.AppTree.Nodes.Add vAppTree(2), tvwChild, vAppTree(1), vAppTree(0), vAppTree(3), vAppTree(4)
        End If
        '---------------------------------------------------------------------------
        
    Loop Until EOF(iFree)
    Close #iFree
    
    'open the Application Entries and add them to frmApplication.AppTree------------
    Open sAppPath$ & "Application.txt" For Input As #iFree
    
    'Restet sDel$.------------------------------------------------------------------
    sDel$ = ("*")
    '-------------------------------------------------------------------------------
    
    Do
    
        Line Input #iFree, sAppTree$
        
            'Use Split Command to extract AppTree items.----------------------------
            vAppTree = Split(sAppTree$, sDel$)
            '-----------------------------------------------------------------------
    
        'Search for lstPath (Path of AppTree Entries)-------------------------------
        Set lstItem = frmApplication.lstPath.ListItems.Add(, , vAppTree(1))
        '---------------------------------------------------------------------------
        
    Loop Until EOF(iFree)
    Close #iFree
    '-------------------------------------------------------------------------------

End Sub

'***********************************************************************************
'Save Application Menu Entries.
'***********************************************************************************
Public Sub subSaveApplicationEntries()

    Dim sDel As String
        sDel$ = "!"
    Dim iFree As Integer
        iFree = FreeFile
    Dim sAppPath As String
        sAppPath$ = mProcFunc.ftnAppPath(App.Path)

    Open sAppPath$ & "Application.txt" For Output As #iFree
    
    'Save AppTree Items using a Deliminated File.-----------------------------------
    For Each nNode In frmApplication.AppTree.Nodes
        
        'If it`s a "root" Item, then Save to Deliminated File.----------------------
        If nNode.FullPath = nNode.Text Then
            Print #iFree, nNode.Text; sDel$; nNode.Key; sDel$; "root"; sDel$; nNode.Image; sDel$; nNode.SelectedImage
        '---------------------------------------------------------------------------
        
        'If it`s a Sub Item, then Save as before but with nNode.Parent.Key.---------
        Else
            Print #iFree, nNode.Text; sDel$; nNode.Key; sDel$; nNode.Parent.Key; sDel$; nNode.Image; sDel$; nNode.SelectedImage
        End If
        '---------------------------------------------------------------------------
    
    Next nNode
    '-------------------------------------------------------------------------------
    
    'Save lstPath (Path of Applications in AppTree) in same File.-------------------
    Dim x As Long
    For x = 1 To frmApplication.lstPath.ListItems.Count
        
        Print #iFree, "*" & frmApplication.lstPath.ListItems(x).Text & "*"
    
    Next x
        
    Close #iFree
    '-------------------------------------------------------------------------------

End Sub


