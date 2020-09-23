Attribute VB_Name = "Module1"
Public Function getcolor(Mode As Integer, Value As Integer) As OLE_COLOR
Select Case Mode
    Case 0
        getcolor = RGB(Value, 0, 0)
    Case 1
        getcolor = RGB(0, Value, 0)
    Case 2
        getcolor = RGB(0, 0, Value)
    Case 3
        getcolor = RGB(Value, 0, Value)
    Case 4
        getcolor = RGB(Value, Value, 0)
    Case 5
        getcolor = RGB(0, Value, Value)
    Case 6
        getcolor = RGB(Value, Value, Value)
End Select
End Function

Function Gettext(Index As Integer) As String
Select Case Index
    Case 0
        Gettext = "Proxy Server Status - Idle"
    Case 1
        Gettext = "Chat Server Status - Idle"
    Case 2
        Gettext = "Cache Server Status - Idle"
    Case 3
        Gettext = "Firewall Status - Idle"
    Case 4
        Gettext = "Firewall Status - Idle"
End Select
End Function
Function GettextA(Index As Integer) As String
ip = LTrim(RTrim(Str(Int(Rnd * 256)))) + "." + LTrim(RTrim(Str(Int(Rnd * 256)))) + "." + LTrim(RTrim(Str(Int(Rnd * 256)))) + "." + LTrim(RTrim(Str(Int(Rnd * 256))))
Port = LTrim(RTrim(Str(Int(Rnd * 256))))
Select Case Index
    Case 0
        GettextA = "Proxy Server Status - Active On " + ip + " port " + Port
    Case 1
        GettextA = "Chat Server Status - Active On " + ip + " port " + Port
    Case 2
        GettextA = "Cache Server Status - Active On " + ip + " port " + Port
    Case 3
        GettextA = "Firewall Status - Active On " + ip + " port " + Port
    Case 4
        GettextA = "Firewall Status - Active On " + ip + " port " + Port
End Select
End Function
