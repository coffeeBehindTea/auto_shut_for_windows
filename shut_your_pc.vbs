' d = CDate("2023/12/16")
' msgbox weekday(d)

function Mut()
    For i = 0 To 20
        Set Ws = Wscript.CreateObject("Wscript.Shell")
        Ws.Sendkeys(Chr(&H88AE))
        Wscript.Sleep(50)
    Next
end function



' if msgbox("start?",vbyesno," ")=vbyes then
if weekday(date) > 1 And weekday(date) < 6 And Hour(time) < 21 then
    
    Mut()
    

' fromday = Date
' toDate="01-Jan-2024 00:00:00"
' toDate = CDate("2024/01/11")

' msgbox "离大考还有" & DateDiff("d", fromday, toDate) & "天~"


    Digital = time
    h = Hour(Digital)
    m = Minute(Digital)
    t = False

    Do Until Second(Digital) = 0
        Digital = time
        t = False
    Loop

    ' dim fso,testfile
    ' set fso = createobject("scripting.filesystemobject") 
    ' set testfile = fso.createtextfile("d:\nice.txt",true)
    ' testfile.writeline("loop started") 
    ' testfile.close

    Do Until h = 22
        Digital = time
        h = Hour(Digital)
        m = Minute(Digital)
        if (h = 8 And m = 0 And t = False) then
            ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 8 And m = 55 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 9 And m = 50 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 10 And m = 45 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 13 And m = 15 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 14 And m = 10 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 15 And m = 10 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 16 And m = 5 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 18 And m = 0 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 18 And m = 55 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        elseif (h = 19 And m = 50 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
            t = True

        else 
            t = False
        end if

        Set Ws = Wscript.CreateObject("Wscript.Shell")
        Wscript.Sleep(1000 * 60)
    Loop

elseif weekday(date) = 6 Then
    Digital = time
    h = Hour(Digital)
    m = Minute(Digital)
    t = False

    Do Until Second(Digital) = 0
        Digital = time
        t = False
    Loop

    Do Until h = 22
        Digital = time
        h = Hour(Digital)
        m = Minute(Digital)
        if (h = 8 And m = 0 And t = False) then
            ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 8 And m = 55 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 9 And m = 50 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 10 And m = 45 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 12 And m = 30 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 13 And m = 25 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True

        elseif (h = 14 And m = 20 And t = False) then
        ' Set Ws = Wscript.CreateObject("Wscript.Shell")
            Mut()
        t = True
        else
            t = False
        end if
        Set Ws = Wscript.CreateObject("Wscript.Shell")
        Wscript.Sleep(1000 * 60)
    Loop

end if
