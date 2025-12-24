Attribute VB_Name = "StaticVariables"
Option Explicit
Option Private Module

Public Const IS_TEST = True

Public json_options As JsonOptionsType

Public progress_bar_next_run_time As Date
Public progress_bar_active As Boolean

Public m_app As AppContext


' Application.OnTimeで使う用 { プロシージャ名: Date }
Public on_time_tasks As Object

