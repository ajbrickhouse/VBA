Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub OpenUrl()

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "https://app.copia.io/DFS/G_L_V_0_P_0.23")

End Sub

