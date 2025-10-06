Attribute VB_Name = "StaticProcedures"
Option Explicit
Option Private Module

Public Function App() As AppContext
    If m_app Is Nothing Then Set m_app = New AppContext
    Set App = m_app
End Function



Public Property Get DBL_QUOTE() As String
    DBL_QUOTE = Chr(34)    ' 「"」を返却します。
End Property

Public Property Get SL_QUOTE() As String
    SL_QUOTE = "'"        ' 「'」を返却します。
End Property


