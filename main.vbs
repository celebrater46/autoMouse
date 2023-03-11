' THIS PROGRAM IS NOT WORKING WITHOUT MS EXCEL
executeGlobal createObject("scripting.FileSystemObject").openTextFile("autoMouse.vbs", 1, false).readAll()

dim myPoint, x, y 
myPoint = apiGetMsgPos
x = myPoint(0)
y = myPoint(1)
msgBox "x = " & x & vbCrLf & "y = " & y 

mouseMove x + 100, y + 100
mouseClickR
msgBox "Clicked!"

mouseMove 1000, 500
mouseClickShift 
msgBox "Shift + Clicked!"

mouseMove 1000, 300
doubleClick
msgBox "Double Clicked!"