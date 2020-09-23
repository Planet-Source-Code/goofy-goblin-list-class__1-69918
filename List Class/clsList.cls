VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim blmListEmpty As Boolean

Dim strList() As String
Dim lngItemData() As Long

Dim lngNewIndex As Long

Dim blmSorted As Boolean
Dim strTag As String

Public Sub AddItem(ByVal strItem As String, Optional ByVal lngIndex As Long = -1)
    Dim i As Integer
    
    If blmListEmpty Then
        lngNewIndex = 0
    Else
        If lngIndex = -1 Then
            lngNewIndex = UBound(strList) + 1
        Else
            If lngIndex < -1 And lngIndex > UBound(strList) Then Exit Sub
            lngNewIndex = lngIndex
        End If
    End If
    
    If blmListEmpty Then
        ReDim Preserve strList(0) As String
        ReDim Preserve lngItemData(0) As Long
        blmListEmpty = False
    Else
        ReDim Preserve strList(UBound(strList) + 1) As String
        ReDim Preserve lngItemData(UBound(strList) + 1) As Long
    End If
    
    If lngNewIndex < UBound(strList) Then
        For i = UBound(strList) To lngNewIndex + 1 Step -1
           strList(i) = strList(i - 1)
           lngItemData(i) = lngItemData(i - 1)
        Next i
    End If
    
    strList(lngNewIndex) = strItem
    lngItemData(lngNewIndex) = 0
    
    If blmSorted Then Call SortList(0, UBound(strList))
    
End Sub

Public Sub Clear()
    
    If blmListEmpty Then Exit Sub
    
    blmListEmpty = True
    
    lngNewIndex = -1
    ReDim strList(0) As String
    ReDim lngItemData(0) As Long
    
End Sub

Public Property Let ItemData(ByVal lngIndex As Long, ByVal lngNewValue As Long)
    
    If lngIndex < 0 Or lngIndex > UBound(strList) Then Exit Property
    
    lngItemData(lngIndex) = lngNewValue
    
End Property

Public Property Get ItemData(ByVal lngIndex As Long) As Long
    
    If lngIndex < 0 Or lngIndex > UBound(strList) Then Exit Property
    
    ItemData = lngItemData(lngIndex)
    
End Property

Public Property Let List(ByVal lngIndex As Long, ByVal strNewValue As String)
    
    If lngIndex < 0 Or lngIndex > UBound(strList) Then Exit Property
    
    strList(lngIndex) = strNewValue
    
    If blmSorted Then Call SortList(0, UBound(strList))
    
End Property

Public Property Get List(ByVal lngIndex As Long) As String
    
    If lngIndex < 0 Or lngIndex > UBound(strList) Then Exit Property
    
    List = strList(lngIndex)
    
End Property

Public Property Get ListCount() As Long
    
    If lngNewIndex = -1 Then Exit Property
    
    ListCount = UBound(strList) + 1
    
End Property

Public Property Get NewIndex() As Long

    NewIndex = lngNewIndex
    
End Property

Public Sub RemoveItem(ByVal lngIndex As Long)
    Dim i As Integer
    
    If blmListEmpty Or lngIndex < 0 Or lngIndex > UBound(strList) Then Exit Sub
    
    For i = lngIndex To UBound(strList) - 1
        strList(i) = strList(i + 1)
        lngItemData(i) = lngItemData(i + 1)
    Next i
    
    If UBound(strList) = 0 Then
        Call Clear
    Else
        ReDim Preserve strList(UBound(strList) - 1) As String
        ReDim Preserve lngItemData(UBound(lngItemData) - 1) As Long
    End If
    
End Sub

Public Property Let Sorted(ByVal blmNewValue As Boolean)
    
    blmSorted = blmNewValue
    
    If blmSorted Then Call SortList(0, UBound(strList))
    
End Property

Public Property Get Sorted() As Boolean
    
    Sorted = blmSorted
    
End Property

Public Property Let Tag(ByVal strNewValue As String)
    
    strTag = strNewValue
    
End Property

Public Property Get Tag() As String
    
    Tag = strTag
    
End Property

Private Sub SortList(ByVal lngLowerBound As Long, ByVal lngUpperBound As Long, Optional ByVal lngCount As Long = 0)
    Dim lngBegin As Long, lngEnd As Long, lngTempLong As Long
    Dim strMiddle As String, strTempString As String
    
    If blmListEmpty Then Exit Sub
    
    lngBegin = lngLowerBound
    lngEnd = lngUpperBound
    strMiddle = strList((lngLowerBound + lngUpperBound) / 2)
    If lngCount = 0 Then _
        lngCount = lngUpperBound - lngLowerBound
    
    Do
        While strList(lngBegin) < strMiddle And lngBegin < lngUpperBound
            lngBegin = lngBegin + 1
        Wend
        While strMiddle < strList(lngEnd) And lngEnd > lngLowerBound
            lngEnd = lngEnd - 1
        Wend
        
        If lngBegin <= lngEnd Then
            strTempString = strList(lngBegin)
            strList(lngBegin) = strList(lngEnd)
            strList(lngEnd) = strTempString
            
            lngTempLong = lngItemData(lngBegin)
            lngItemData(lngBegin) = lngItemData(lngEnd)
            lngItemData(lngEnd) = lngTempLong
            
            lngBegin = lngBegin + 1
            lngEnd = lngEnd - 1
        End If
        
    Loop While lngBegin <= lngEnd
    
    If lngLowerBound < lngEnd Then SortList lngLowerBound, lngEnd, lngCount
    If lngBegin < lngUpperBound Then SortList lngBegin, lngUpperBound, lngCount
    
End Sub

Private Sub Class_Initialize()
    
    blmListEmpty = True
    
End Sub

Private Sub Class_Terminate()
    
    Call Clear
    
End Sub
