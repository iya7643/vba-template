Attribute VB_Name = "StaticVariables"
Option Explicit
Option Private Module

Public Const IS_TEST = True

Public json_options As JsonOptionsType

Public progress_bar_next_run_time As Date
Public progress_bar_active As Boolean

Public m_app As AppContext


' Application.OnTimeで使う用（dict[プロシージャ名: Date]）
Public on_time_tasks As Object

' AppEventHandlerで使う用（{ ブックのパス|シート名: { セルアドレス: { "VALUE": 値, "PROC": 実行するプロシージャ名 } } }）
Public watch_items As Object
