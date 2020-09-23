VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Class"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEdit 
      Caption         =   "Edit"
      Height          =   2775
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Item"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox chkSorted 
         Caption         =   "Sorted"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame fraPreview 
      Caption         =   "Preview"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstPreview 
         Height          =   2400
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objList As New clsList

Private Sub RefreshPreview()
    Dim i As Integer
    
    If objList.ListCount = 0 Then Exit Sub
    
    lstPreview.Clear
    For i = 0 To objList.ListCount - 1
        lstPreview.AddItem objList.List(i)
        lstPreview.ItemData(lstPreview.NewIndex) = objList.ItemData(i)
    Next i
    
End Sub

Private Sub chkSorted_Click()
    
    objList.Sorted = CBool(chkSorted.Value)
    
    Call RefreshPreview
    
End Sub

Private Sub cmdAddItem_Click()
    Dim strUserInput As String
    
    strUserInput = InputBox("Please enter the text for your new list item.", "Enter Text")
    If LenB(strUserInput) = 0 Then Exit Sub
    
    objList.AddItem strUserInput
    
    Call RefreshPreview
    
End Sub

Private Sub cmdClear_Click()
    
    objList.Clear
    
    Call RefreshPreview
    
End Sub

Private Sub cmdRemoveItem_Click()
    
    If lstPreview.ListIndex < 0 Then Exit Sub
    
    objList.RemoveItem lstPreview.ListIndex
    
    Call RefreshPreview
    
End Sub

Private Sub lstPreview_Click()
    
    If lstPreview.ListIndex < 0 Then Exit Sub
    
    lstPreview.ToolTipText = CStr(lstPreview.ItemData(lstPreview.ListIndex))
    
End Sub
