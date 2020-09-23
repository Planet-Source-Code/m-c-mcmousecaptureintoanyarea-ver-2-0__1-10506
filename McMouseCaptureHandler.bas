Attribute VB_Name = "McMouseCaptureHandler"
'Do not erase or alter these comments !
'Module Name: McMouseCaptureHandler
'version 2.0 (August,2000)
'Author: Miran Cvenkel
'Changes from v. 1: MCCaptureMouseCursorIntoSpecialArea function added

'THERE ARE THREE FUNCTIONS, actualy 4 but IsThereMenu works only within this module
'+ some stuf which you should ignore as it is in development !
'**************************************************************
'First function name: MCCaptureMouseCursorIntoArea
'**************************************************************
'Functionality:
'1.Limit mouse movement into rectangle determined by any rectangular VB control
'(could be visible = false!) including image,shape and label that don't have hwnd
'which is usualy used for Limiting mouse moving(it doesn't work on line because
'it is not rectangular but you could make shape so thin that result would be the desired
'one, only in horizontal and vertical position ofcourse)
'2.Handles all border styles,menu, caption bars - which influence resulting position
'3.Handles twip and pixel scale modes
'-----------------USAGE !-----------------------------------
'How to use it?
'first put some controls on form !
'call like this  : MCCaptureMouseCursorIntoArea FormName,ControlNameOnThatForm
'Example  1   : MCCaptureMouseCursorIntoArea Form1,shape1
'Example  2   : MCCaptureMouseCursorIntoArea Form2,image3
'Example  3   : MCCaptureMouseCursorIntoArea Form2,command1
'etc
'that's it, simple isn't it
'IMPORTANT only the folowing unlocks mouse: ClipCursor ByVal 0&
'if you made mistake it helps if u use folowing steps:
'press CTRL + ALT + DEL and then click Cancel - the mouse will be released !
'---------------------------------------------------------------
'**************************************************************************
'Second function name: MCCaptureMouseCursorIntoNestedArea
'**************************************************************************
'Functionality:
'1.extends functionality of MCCaptureMouseCursorIntoArea
'to the controls inside other controls i.e. if u have control nested inside picbox or frame
'-----------------USAGE !-----------------------------------
'call like this  : MCCaptureMouseCursorIntoNestedArea FormName,MotherControlName,NetsedControlName
'it is obvious that mother control could be only frame or picbox, nested control can be
'any rectangular control
'Example  1   : MCCaptureMouseCursorIntoNestedArea Form1,picture1,picture2
'Example  2   : MCCaptureMouseCursorIntoArea Form2,frame1,list1
'Example  3   : MCCaptureMouseCursorIntoArea Form2,picture3,shape1
'etc
'***********************************************************************
'Third function name: MCCaptureMouseCursorIntoSpecialArea
'***********************************************************************
'Functionality:
'Capture mouse into diferent forms areas like: CaptionBar, CompleteForm,
'CompleteFormBorderExcluded , FormClientArea
'-----------------USAGE !-----------------------------------
'call like this  :
'MCCaptureMouseCursorIntoSpecialArea(SourceForm As Form, Area As String)
'Area can be: "CaptionBar" or "FormClientArea" or  "CompleteForm" or
' "CompleteFormBorderExcluded"
'the last one is useful if u want sizable border and  want to disable user to size
'form with mouse
'Example: 'MCCaptureMouseCursorIntoSpecialArea(Me, "CompleteForm" )
  
Private Const SM_CYCAPTION = 4  'Height of windows caption
Private Const SM_CXBORDER = 5  'Width of no-sizable borders
Private Const SM_CYBORDER = 6  'Height of non-sizable borders
Private Const SM_CXDLGFRAME = 7  'Width of dialog box borders
Private Const SM_CYDLGFRAME = 8  'Height of dialog box borders
Private Const SM_CYMENU = 15  'Height of menu
Private Const SM_CXFRAME = 32 'normal borders width - sizable
Private Const SM_CYFRAME = 33 'normal borders height - sizable
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Type POINT
    X As Long
    Y As Long
End Type
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'this is used only to determine if there is a visible menu on our form
'and in MCCaptureMouseCursorIntoSpecialArea
Private Declare Sub GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINT)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private IsThereMenuAnswer As Boolean
'this one must be public to enable you to unlock mouse capture
Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Function MCCaptureMouseCursorIntoArea(SourceForm As Form, SourceControl As Control)

Dim MenuHeight As Byte ' if it is there
Dim BorderWidth As Byte ' if there is any border
Dim BorderHeight As Byte ' if there is any border
Dim CaptionBarHeight As Byte ' if there is any caption bar
          
    '1.   First ask ourself what kind of form ,that our control is on, we have ?
    '      Does it have menu, what kind of border does it have - all this influence
    '      where our mouse controling rectangle will come up
                            
'+++++++++++++++
'end app if fatal error, which occurs if we are dealing with picbox and it have diferent scale mode as form
If TypeOf SourceControl Is PictureBox Then
         If SourceForm.ScaleMode <> SourceControl.ScaleMode Then
         MsgBox "SourceForm and SourceControl must both have same scalemode - if exist, (twip or pixel). This is not so now .... ending, ", vbCritical, "MCCaptureMouseCursorIntoArea Function Message"
         End
         End If
End If
'+++++++++++++++
                      
    'what kind of border doe's it (form) have ?
    'accordingly get & calculate system metrics for form borders
    Select Case SourceForm.BorderStyle
               Case 0 'none
               BorderWidth = 0
               BorderHeight = 0
               CaptionBarHeight = 0
               'special case - you  have border style = 0 and menu on form
                        Dim ctl As Control
                        For Each ctl In SourceForm
                        If TypeOf ctl Is Menu Then
                            BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
                            BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
                            CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
                        End If
                        Next
               Case 1
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 2 'sizable
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 3 'dialog
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 4 'fixed
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CYDLGFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case 5
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case Else
    End Select
      
    
    'is there a menu on our source form ? if it is - get it's height
    IsThereMenu SourceForm, BorderHeight, CaptionBarHeight 'function is on the bottom of module
    If IsThereMenuAnswer = True Then
    MenuHeight = GetSystemMetrics(SM_CYMENU)
    Else: MenuHeight = 0
    End If
     
   Dim TargetArea As RECT 'this will eat data about the size of area
                                         'where captured mouse should be moving
   
   '2. Fill TargetArea rect  with data - now it awares what size it is & place
   '    it where it should be
   
   'following code handles twips or pixels (form.scalemode)
   Select Case SourceForm.ScaleMode
   Case 1 'twip
       'Get information about our control (which determines area to capture mouse in)
        TargetArea.top = SourceControl.top / Screen.TwipsPerPixelY
        TargetArea.left = SourceControl.left / Screen.TwipsPerPixelX
        TargetArea.right = (SourceControl.left + SourceControl.Width) / Screen.TwipsPerPixelX
        TargetArea.bottom = (SourceControl.top + SourceControl.Height) / Screen.TwipsPerPixelY
        'screen coordinates to move TargetArea rectangle to
        xmove = BorderWidth + (SourceForm.left / Screen.TwipsPerPixelX) - TargetArea.left + (SourceControl.left / Screen.TwipsPerPixelX)
        ymove = BorderHeight + CaptionBarHeight + MenuHeight + (SourceForm.top / Screen.TwipsPerPixelY) - TargetArea.top + (SourceControl.top / Screen.TwipsPerPixelY)
   Case 3 'Pixel
         'Get information about our control (which determines area to capture mouse in)
        TargetArea.top = SourceControl.top
        TargetArea.left = SourceControl.left
        TargetArea.right = SourceControl.left + SourceControl.Width
        TargetArea.bottom = SourceControl.top + SourceControl.Height
        'screen coordinates to move TargetArea rectangle to
        xmove = BorderWidth + (SourceForm.left / Screen.TwipsPerPixelX) - TargetArea.left + (SourceControl.left)
        ymove = BorderHeight + CaptionBarHeight + MenuHeight + (SourceForm.top / Screen.TwipsPerPixelX) - TargetArea.top + (SourceControl.top)
   Case Else
        MsgBox "Only twip or pixel scale mode allowed", vbCritical, "MCCaptureMouseCursorIntoArea Function Message"
        End
   End Select
    
    'now actualy move it there, I mean TargetArea rectangle
    'it is like it would invisibly (transculent) cover sourcecontrol
    OffsetRect TargetArea, xmove, ymove
    
    'limit the cursor movement into that rect
    ClipCursor TargetArea
    
    'huh, this was a lot of work, lol
 End Function
Public Function MCCaptureMouseCursorIntoNestedArea(SourceForm As Form, MotherControl As Control, NestedControl As Control)

Dim MenuHeight As Byte ' if it is there
Dim BorderWidth As Byte ' if there is any border
Dim BorderHeight As Byte ' if there is any border
Dim CaptionBarHeight As Byte ' if there is any caption bar
          
    '1.   First ask ourself what kind of form ,that our control is on, we have ?
    '      Does it have menu, what kind of border does it have - all this influence
    '      where our mouse controling rectangle will come up
    
    
                                                
    'what kind of border doe's it (form) have ?
    'accordingly get & calculate system metrics for form borders
    Select Case SourceForm.BorderStyle
               Case 0 'none
               BorderWidth = 0
               BorderHeight = 0
               CaptionBarHeight = 0
               'special case - you  have border style = 0 and menu on form
                        Dim ctl As Control
                        For Each ctl In SourceForm
                        If TypeOf ctl Is Menu Then
                            BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
                            BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
                            CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
                        End If
                        Next
               Case 1
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 2 'sizable
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 3 'dialog
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 4 'fixed
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CYDLGFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case 5
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case Else
    End Select
    
       'is there a menu on our source form ? if it is - get it's height
    IsThereMenu SourceForm, BorderHeight, CaptionBarHeight 'function is on the bottom of module
    If IsThereMenuAnswer = True Then
    MenuHeight = GetSystemMetrics(SM_CYMENU)
    Else: MenuHeight = 0
    End If
    
                        
'so far everything the same as in previous function, now diferences
 
'+++++++++++++++
'end if fatal error which occurs if we are dealing with picboxes and they have diferent scale mode as form
If TypeOf MotherControl Is PictureBox And TypeOf NestedControl Is PictureBox Then
             If SourceForm.ScaleMode <> MotherControl.ScaleMode Or _
             SourceForm.ScaleMode <> NestedControl.ScaleMode Or _
             NestedControl.ScaleMode <> MotherControl.ScaleMode Then
             MsgBox "SourceForm, MotherControl and NestedControl must all have same scalemode - if exist, (twip or pixel). This is not so now .... ending, ", vbCritical, "MCCaptureMouseCursorIntoNestedArea Function Message"
             End
             End If
End If

If TypeOf MotherControl Is PictureBox And SourceForm.ScaleMode <> MotherControl.ScaleMode Then
MsgBox "SourceForm, MotherControl and NestedControl must all have same scalemode - if exist, (twip or pixel). This is not so now .... ending, ", vbCritical, "MCCaptureMouseCursorIntoNestedArea Function Message"
End
End If

If TypeOf NestedControl Is PictureBox And SourceForm.ScaleMode <> MotherControl.ScaleMode Then
MsgBox "SourceForm, MotherControl and NestedControl must all have same scalemode - if exist, (twip or pixel). This is not so now .... ending, ", vbCritical, "MCCaptureMouseCursorIntoNestedArea Function Message"
End
End If
'+++++++++++++++


'ask if mothercontrol (it could be only frame or picturebox) have border ?
If MotherControl.BorderStyle = 1 Then ' if it has one
AdjustmentCausedByBorder = 2 ' we need to repair coordinates where rect will go
Else: AdjustmentCausedByBorder = 0
End If


Dim NestedTargetArea As RECT 'this will eat data about the size of area
                                                 'where captured mouse should be moving
   
   '2. Fill TargetArea rect  with data - now it awares what size it is & place
   '    it where it should be
   
        Select Case SourceForm.ScaleMode
        Case 1 'twip
       'Get information about our control (which determines area to capture mouse in)
        NestedTargetArea.top = (MotherControl.top + NestedControl.top) / Screen.TwipsPerPixelY
        NestedTargetArea.left = (MotherControl.left + NestedControl.left) / Screen.TwipsPerPixelX
        NestedTargetArea.right = (MotherControl.left + NestedControl.left + NestedControl.Width) / Screen.TwipsPerPixelX
        NestedTargetArea.bottom = (MotherControl.top + NestedControl.top + NestedControl.Height) / Screen.TwipsPerPixelY
        'screen coordinates to move TargetArea rectangle to
        xmove = BorderWidth + (SourceForm.left / Screen.TwipsPerPixelX) - NestedTargetArea.left / Screen.TwipsPerPixelY + (NestedTargetArea.left / Screen.TwipsPerPixelX) + AdjustmentCausedByBorder
        ymove = BorderHeight + CaptionBarHeight + MenuHeight + (SourceForm.top / Screen.TwipsPerPixelY) - NestedTargetArea.top / Screen.TwipsPerPixelY + (NestedTargetArea.top / Screen.TwipsPerPixelY) + AdjustmentCausedByBorder
        Case 3 'pixel
         'Get information about our control (which determines area to capture mouse in)
        NestedTargetArea.top = MotherControl.top + NestedControl.top
        NestedTargetArea.left = MotherControl.left + NestedControl.left
        NestedTargetArea.right = MotherControl.left + NestedControl.left + NestedControl.Width
        NestedTargetArea.bottom = MotherControl.top + NestedControl.top + NestedControl.Height
        'screen coordinates to move TargetArea rectangle to
        xmove = BorderWidth + (SourceForm.left / Screen.TwipsPerPixelX) - NestedTargetArea.left + (NestedTargetArea.left) + AdjustmentCausedByBorder
        ymove = BorderHeight + CaptionBarHeight + MenuHeight + (SourceForm.top / Screen.TwipsPerPixelX) - NestedTargetArea.top + (NestedTargetArea.top) + AdjustmentCausedByBorder
        Case Else
        MsgBox "Only twip or pixel scale mode allowed", vbCritical, "MCCaptureMouseCursorIntoNestedArea Function Message"
        End
        End Select
    
    'now actualy move it there, I mean TargetArea rectangle
    'it is like it would invisibly(transculent) cover sourcecontrol
    OffsetRect NestedTargetArea, xmove, ymove
    
    'limit the cursor movement into that rect
    ClipCursor NestedTargetArea
End Function
Public Function MCCaptureMouseCursorIntoSpecialArea(SourceForm As Form, Area As String)
Dim topleft As POINT
 'what kind of border doe's it (form) have ?
    'accordingly get & calculate system metrics for form borders
    Select Case SourceForm.BorderStyle
               Case 0 'none
               BorderWidth = 0
               BorderHeight = 0
               CaptionBarHeight = 0
               'special case - you  have border style = 0 and menu on form
                        Dim ctl As Control
                        For Each ctl In SourceForm
                        If TypeOf ctl Is Menu Then
                            BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
                            BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
                            CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
                        End If
                        Next
               Case 1
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 2 'sizable
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 3 'dialog
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION)
               Case 4 'fixed
               BorderWidth = GetSystemMetrics(SM_CXDLGFRAME)
               BorderHeight = GetSystemMetrics(SM_CYDLGFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case 5
               BorderWidth = GetSystemMetrics(SM_CXFRAME)
               BorderHeight = GetSystemMetrics(SM_CYFRAME)
               'seems to me it is smaller then ordinary one
               CaptionBarHeight = GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYDLGFRAME)
               Case Else
    End Select


'Now action
Select Case Area
Case "MenuBar"

           'is there a menu on our source form ? if it is - get it's height
         '  IsThereMenu SourceForm, BorderHeight, CaptionBarHeight 'function is on the bottom of module
         '  If IsThereMenuAnswer = True Then
         '  MenuHeight = GetSystemMetrics(SM_CYMENU)
         '  Else: MenuHeight = 0: MsgBox "There is no menu on this form - can't capture mouse - Exit"
         '  End If

Case "CaptionBar"
          
           
                 Dim MenuRect As RECT
                 Dim MiscRect As RECT
                 'get our form coordinates
                 GetWindowRect SourceForm.hwnd, MenuRect  'get our form measures
                 GetClientRect SourceForm.hwnd, MiscRect
                 'Adjust
                 MenuRect.top = MenuRect.top + BorderHeight
                 MenuRect.bottom = MenuRect.bottom - BorderHeight - (MiscRect.bottom - MiscRect.top)
                 MenuRect.left = MenuRect.left + BorderWidth
                 MenuRect.right = MenuRect.right - BorderWidth
                 ClipCursor MenuRect
           
           
Case "FormClientArea"
                 Dim SourceFormClientRect As RECT
                 'get our form coordinates
                 GetClientRect SourceForm.hwnd, SourceFormClientRect 'get our form "working area" measures
                 'Convert window coordinates to screen coordinates
                 topleft.Y = SourceFormClientRect.top
                 topleft.X = SourceFormClientRect.left
                 ClientToScreen SourceForm.hwnd, topleft
                 'move rect there
                 OffsetRect SourceFormClientRect, topleft.X, topleft.Y
                 ClipCursor SourceFormClientRect
Case "CompleteFormBorderExcluded"
                 Dim SourceFormRectBEx As RECT
                 'get our form coordinates
                 GetWindowRect SourceForm.hwnd, SourceFormRectBEx  'get our form measures
                 SourceFormRectBEx.top = SourceFormRectBEx.top + BorderHeight
                 SourceFormRectBEx.bottom = SourceFormRectBEx.bottom - BorderHeight
                 SourceFormRectBEx.left = SourceFormRectBEx.left + BorderWidth
                 SourceFormRectBEx.right = SourceFormRectBEx.right - BorderWidth
                 ClipCursor SourceFormRectBEx
Case "CompleteForm"
                 Dim SourceFormRect As RECT
                 'get our form coordinates
                 GetWindowRect SourceForm.hwnd, SourceFormRect  'get our form measures
                 ClipCursor SourceFormRect
Case Else
End Select

End Function
 
 
 Private Function IsThereMenu(SourceForm As Form, BorderHeight As Byte, CaptionBarHeight As Byte)
 'After many different attempts, this proved to be the only reliable way
 'to tell if there is a visible menu on form or not
 
 Dim SourceFormRect As RECT
 GetClientRect SourceForm.hwnd, SourceFormRect 'get our form "working area" measures
 'now calculate diference between SourceForm.Height(which is the size of
 'rectangle determining whole form) and (clientarea height+ (2*borderheight) + captionbarheight)
 'if diference = SystemMenuHeight then it must be there, else not
   
   Select Case SourceForm.ScaleMode
   Case 1 'twip
   f = (((SourceForm.Height - ((SourceFormRect.bottom - SourceFormRect.top) * Screen.TwipsPerPixelY)) / Screen.TwipsPerPixelY) - (2 * BorderHeight)) - CaptionBarHeight
   Beep
   Case 3 'pixel
  f = (((SourceForm.Height / Screen.TwipsPerPixelY) - (SourceFormRect.bottom - SourceFormRect.top)) - (2 * BorderHeight)) - CaptionBarHeight
   Beep
   Case Else
  MsgBox "ERROR - only twip and pixel scale mode allowed for this function", vbCritical, "McMouseCaptureHandler module"
   End Select
 
  'get answer
  If f = GetSystemMetrics(SM_CYMENU) Then
  IsThereMenuAnswer = True
  Else: IsThereMenuAnswer = False
  End If
  
 End Function
Public Function PopUpMenuStuff(WinHwnd As Integer)
 Dim SourceFormRect As RECT
                 'get our form coordinates
                 GetWindowRect WinHwnd, SourceFormRect  'get our form measures
                 ClipCursor SourceFormRect
 End Function
