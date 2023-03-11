option explicit 
dim x, y, objWsh
set excel = wscript.createObject("excel.application")

' mouse constants
const mouseEventF_absolute = &H8000
const mouseEventF_absolute = 32768
const mouseMove = &H1
const mouseEventF_leftDown = &H2
const mouseEventF_leftUp = &H4

moveMouse 960, 540
mouseClickR

sub mouseClickR
    dim dwFlags : dwFlags = mouseEventF_leftDown or mouseEventF_leftUp
    call apiMouseEvent(dwFlags, 0, 0, 0, 0)
    wscript.sleep 100
end sub

sub moveMouse(x, y)
    dim mx, my, dwFlags
    const screenX = 1024
    const screenY = 768
    dwFlags = mouseEventF_absolute + mouseMove 
    mx = int(x * 65535 / screenX)
    my = int(y * 65535 / screenY)
    call apiMouseEvent(dwFlags, mx, my, 0, 0)
    wscript.sleep 100
end sub

sub apiMouseEvent(dwFlags, dx, dy, dwData, dwExtraInfo)
    const apiStr = "call(""user32"",""mouse_event"",""JJJJJJ"", $1, $2, $3, $4, $5)"
    dim strFunc : strFunc = replace(replace(replace(replace(replace(apiStr, "$1", dwFlags), "$2", dx), "$3", dy), "$4", dwData), "$5", dwExtraInfo)
    call excel.executeExcel4macro(strFunc)
end sub
' Option Explicit

' Dim x, y
' Dim Excel

' 'シェルオブジェクトの作成
' Set Excel = WScript.CreateObject("Excel.Application")

' 'キーコード
' 'Const VK_SHIFT = &H10

' 'マウス定数
' Const MOUSEEVENTF_ABSOLUTE = &H8000
' Const MOUSEEVENTF_ABSOLUTE = 32768
' Const MOUSE_MOVE = &H1
' Const MOUSEEVENTF_LEFTDOWN = &H2
' Const MOUSEEVENTF_LEFTUP = &H4


' MouseMove 700 , 250　　'←ここで移動させたい座標を指定する
' MouseClick


' 'クリック
' Sub MouseClick
' 　Dim dwFlags
' 　dwFlags = MOUSEEVENTF_LEFTDOWN or MOUSEEVENTF_LEFTUP
' 　Call API_mouse_event(dwFlags, 0, 0, 0, 0)
' 　WScript.Sleep 100
' End Sub



' 'マウスポインタ移動
' Sub MouseMove(x, y)
' 　Dim pos_x, pos_y, dwFlags
' 　Const SCREEN_X = 1024
' 　Const SCREEN_Y = 768

' 　dwFlags = MOUSEEVENTF_ABSOLUTE + MOUSE_MOVE
' 　pos_x = Int(x * 65535 / SCREEN_X)
' 　pos_y = Int(y * 65535 / SCREEN_Y)
' 　Call API_mouse_event(dwFlags, pos_x, pos_y, 0, 0)
' 　WScript.Sleep 100
' End Sub



' 'APIを叩く
' Sub API_mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo)
' 　Dim strFunction
' 　Const API_STRING = "CALL(""user32"",""mouse_event"",""JJJJJJ"", $1, $2, $3, $4, $5)"
' 　strFunction = Replace(Replace(Replace(Replace(Replace(API_STRING, "$1", dwFlags), "$2", dx), "$3", dy), "$4", dwData), "$5", dwExtraInfo)
' 　Call Excel.ExecuteExcel4Macro(strFunction)
' End Sub