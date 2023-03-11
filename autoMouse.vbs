option explicit 
dim x, y, objWsh
set objWsh = wscript.createObject("excel.application")

' mouse constants
const mouseEventF_absolute = &H8000
' const mouseEventF_absolute = 32768
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

sub apiKeyEvent(bVk, bScan, dwFlags, dwExtraInfo)
    const apiStr = "call(""user32"",""keyEvent"",""JJJJJ"", $1, $2, $3, $4)"
    dim strFunc : strFunc = replace(replace(replace(replace(apiStr, "$1", bVk), "$2", bScan), "$3", dwFlags), "$4", dwExtraInfo)
    call excel.executeExcel4macro(strFunc)
    strFunc = replace
end sub

function apiGetMsgPos()
    dim ret, strHex, x, y, strFunc
    const apiStr = "call(""user32"",""getMsgPos"",""J"")"
    strFunc = apiStr
    ret = excel.executeExcel4Macro(strFunc)
    strHex = right("00000000" & hex(ret), 8)
    x = clng("&H" & right(strHex, 4))
    y = clng("&H" & left(strHex, 4))
    apiGetMsgPos = array(x, y)
end function
