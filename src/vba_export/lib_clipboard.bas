Attribute VB_Name = "lib_clipboard"
Option Explicit

Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "User32" () As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal wFormat As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr


Public Sub SetText(Text As String)
'***************************************************************************
'Purpose: puts the text from a string object onto the clipboard
'Inputs
'Outputs:
'***************************************************************************
    #If VBA7 Then
        Dim hGlobalMemory As LongPtr
        Dim lpGlobalMemory As LongPtr
        Dim hClipMemory As LongPtr
    #Else
        Dim hGlobalMemory As Long
        Dim lpGlobalMemory As Long
        Dim hClipMemory As Long
    #End If
    
    Const GHND = &H42
    Const CF_TEXT = 1
    
    ' Allocate moveable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GHND, Len(Text) + 1)
    
    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    
    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)
    
    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could Not unlock memory location. Copy aborted."
        GoTo CloseClipboard
    End If
    
    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could Not open the Clipboard. Copy aborted."
        Exit Sub
    End If
    
    ' Clear the Clipboard.
    Call EmptyClipboard
    
    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
    
CloseClipboard:
    If CloseClipboard() = 0 Then
        MsgBox "Could Not close Clipboard."
    End If
    
End Sub


Public Property Get GetText()
'***************************************************************************
'Purpose: gets the text from the clipboard into a string object
'Inputs
'Outputs:
'***************************************************************************
    #If VBA7 Then
        
        Dim hClipMemory As LongPtr
        Dim lpClipMemory As LongPtr
        
    #Else
        
        Dim hClipMemory As Long
        Dim lpClipMemory As Long
        
    #End If
    
    Dim MaximumSize As Long
    Dim ClipText    As String
    
    Const CF_TEXT = 1
    
    If OpenClipboard(0&) = 0 Then
        MsgBox "Cannot open Clipboard. Another app. may have it open"
        Exit Property
    End If
    
    ' Obtain the handle to the global memory block that is referencing the text.
    hClipMemory = GetClipboardData(CF_TEXT)
    If IsNull(hClipMemory) Then
        MsgBox "Could Not allocate memory"
        GoTo CloseClipboard
    End If
    
    ' Lock Clipboard memory so we can reference the actual data string.
    lpClipMemory = GlobalLock(hClipMemory)
    
    If Not IsNull(lpClipMemory) Then
        MaximumSize = 64
        
        Do
            MaximumSize = MaximumSize * 2
            
            ClipText = Space$(MaximumSize)
            Call lstrcpy(ClipText, lpClipMemory)
            Call GlobalUnlock(hClipMemory)
            
        Loop Until ClipText Like "*" & vbNullChar & "*"
        
        ' Peel off the null terminating character.
        ClipText = Left$(ClipText, InStrRev(ClipText, vbNullChar) - 1)
        
    Else
        MsgBox "Could Not lock memory To copy String from."
    End If
    
CloseClipboard:
    Call CloseClipboard
    GetText = ClipText
    
End Property
