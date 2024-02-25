Sub searchThis()
' Viraj: Search the selected Cell contents

  Dim chromePath As String
  Dim srchtxt As String
  Dim search_string As String

 ' Check the path of the Chrome/other browser's executable file, on your machine and update it accordingly

'chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
  chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

' Highlighting the active cell to confirm the search was done on that cell
    With ActiveCell
        .Interior.ColorIndex = 34 ' Turquoise
        srchtxt = .Value
    End With

    search_string = Replace(srchtxt, "&", " ")
    Do While InStr(1, search_string, "  ")
        search_string = Replace(search_string, "  ", " ")
    Loop
    search_string = Replace(search_string, " ", "+") ' Remove spaces so that it becomes one single string


' Call Chrome and search the string
' Note: This opens up a new Tab\Window when you do the search
'  Shell (chromePath & " -url http://www.google.com/search?q=" & search_string)

   Shell (chromePath & " -url https://www.google.com/search?&pws=0&gl=us&gws_rd=cr&q=" & search_string)

End Sub

