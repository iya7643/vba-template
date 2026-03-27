VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_SLACK = 0
#Const USE_GOOGLE_CHAT = 0
#Const USE_HOLIDAYS_JP = 0

Private m_json_helper As JsonHelper
Private m_constants As Constants
Private m_utility As Utility
Private m_app_event_handler As AppEventHandler

#If USE_SLACK Then
Private m_api_slack As ApiSlack
#End If

#If USE_GOOGLE_CHAT Then
Private m_google_chat As ApiGoogleChat
#End If

#If USE_HOLIDAYS_JP Then
Private m_holidays_jp As ApiHolidaysJp
#End If

Private m_progress_bar_controller As ProgressBarController


Public Property Get JsonHelper() As JsonHelper
    If m_json_helper Is Nothing Then Set m_json_helper = New JsonHelper
    Set JsonHelper = m_json_helper
End Property

Public Property Get Constants() As Constants
    If m_constants Is Nothing Then
        Set m_constants = New Constants
        Call m_constants.init(JsonHelper)
    End If
    Set Constants = m_constants
End Property

Public Property Get Utility() As Utility
    If m_utility Is Nothing Then Set m_utility = New Utility
    Set Utility = m_utility
End Property

Public Property Get AppEventHandler() As AppEventHandler
    If m_app_event_handler Is Nothing Then Set m_app_event_handler = New AppEventHandler
    Set AppEventHandler = m_app_event_handler
End Property

#If USE_SLACK Then
Public Property Get ApiSlack() As ApiSlack
    If m_api_slack Is Nothing Then Set m_api_slack = New ApiSlack
    Set ApiSlack = m_api_slack
End Property
#End If

#If USE_GOOGLE_CHAT Then
Public Property Get ApiGoogleChat() As ApiGoogleChat
    If m_google_chat Is Nothing Then Set m_google_chat = New ApiGoogleChat
    Set ApiGoogleChat = m_google_chat
End Property
#End If

#If USE_HOLIDAYS_JP Then
Public Property Get ApiHolidaysJp() As ApiHolidaysJp
    If m_holidays_jp Is Nothing Then Set m_holidays_jp = New ApiHolidaysJp
    Set ApiHolidaysJp = m_holidays_jp
End Property
#End If

Public Property Get ProgressBarController() As ProgressBarController
    If m_progress_bar_controller Is Nothing Then Set m_progress_bar_controller = New ProgressBarController
    Set ProgressBarController = m_progress_bar_controller
End Property

