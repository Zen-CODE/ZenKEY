Attribute VB_Name = "modBalloon"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Variable to hold the ID of the
'default window message processing
'procedure. Returned by SetWindowLong
Public defWindowProc As Long

'Get/SetWindowLong messages
Public Const GWL_WNDPROC As Long = (-4)
Private Const GWL_HWNDPARENT As Long = (-8)
Private Const GWL_ID As Long = (-12)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_USERDATA As Long = (-21)

'general windows messages
Private Const WM_USER As Long = &H400
Private Const WM_NOTIFY As Long = &H4E
Private Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10

'our app-specific message to trap
'in the WindowProc routine
Private Const WM_APP As Long = &H8000&
Public Const WM_MYHOOK As Long = WM_APP + &H15

'mouse constants for the callback
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203

Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209

Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206


'this identifies this icon, allowing an app
'to display more than one (each with unique IDs)
'and respond to each individually.
Public Const APP_SYSTRAY_ID = 999

Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)


Private Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hwnd As Long) As Long
   
Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Any) As Long

Public Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" _
   (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long



Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

  'If the handle returned is to our form,
  'call a message handler to deal with
  'tray notifications. If it is a general
  'system message, pass it on to
  'the default window procedure.
  '
  'If destined for the form and equal to
  'our custom hook message (WM_MYHOOK),
  'examining lParam reveals the message
  'generated, to which we react appropriately.
   On Error Resume Next
  
   Select Case hwnd
   
     'form-specific handler
      Case Form1.hwnd
         
         Select Case uMsg
          'check uMsg for the application-defined
          'identifier (NID.uID) assigned to the
          'systray icon in NOTIFYICONDATA (NID).
  
           'WM_MYHOOK was defined as the message sent
           'as the .uCallbackMessage member of
           'NOTIFYICONDATA the systray icon
            Case WM_MYHOOK
            
              'lParam is the value of the message
              'that generated the tray notification.
               Select Case lParam
                  Case WM_RBUTTONUP:

                 'This assures that focus is restored to
                 'the form when the menu is closed. If the
                 'form is hidden, it (correctly) has no effect.
                  Call SetForegroundWindow(Form1.hwnd)
                  Form1.PopupMenu Form1.zmnuSysTrayDemo
               
                  Case NIN_BALLOONSHOW
                     Debug.Print "The balloon tip has just appeared"
                            
                  Case NIN_BALLOONHIDE
                     Debug.Print "The systray icon was removed when" & _
                            "the balloon tip was displayed"
                            
                  Case NIN_BALLOONUSERCLICK
                     Debug.Print "The user clicked on the balloon tip"
                            
                  Case NIN_BALLOONTIMEOUT
                     Debug.Print "The balloon tip timed-out without " & _
                            "the user clicking it, or the user " & _
                            "clicked the balloon tip's close button"
                            
               End Select
            
           'handle any other form messages by
           'passing to the default message proc
            Case Else
            
               WindowProc = CallWindowProc(defWindowProc, _
                                            hwnd, _
                                            uMsg, _
                                            wParam, _
                                            lParam)
               Exit Function
            
         End Select
     
     'this takes care of messages when the
     'handle specified is not that of the form
      Case Else
      
          WindowProc = CallWindowProc(defWindowProc, _
                                      hwnd, _
                                      uMsg, _
                                      wParam, _
                                      lParam)
   End Select
   
End Function
