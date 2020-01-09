Option Explicit

DIM intRandom,strSelection

Randomize
intRandom = Int(100*Rnd)


Select Case True
  Case intRandom < 7
    strSelection =  "Jacalito's / Bozelli's"

  Case intRandom < 11
    strSelection =  "Subway"

  Case intRandom < 21
    strSelection =  "Chipotle's"

  Case intRandom < 29
    strSelection =  "Kajana's"

  Case intRandom < 33
    strSelection =  "Saratoga Pizza's"

  Case intRandom < 39
    strSelection =  "Semi Perfect Pita"

  Case intRandom < 49
    strSelection =  "Turds in the window"

  Case intRandom < 57
    strSelection =  "Pho"

  Case intRandom < 62
    strSelection =  "Chick-Fil-A"

  Case intRandom < 68
    strSelection =  "Costco"

  Case intRandom < 69
    strSelection =  "Formerly WaBa Grill"

  Case intRandom < 72
    strSelection =  "BGR Burger"

  Case intRandom < 75
    strSelection =  "5 Guys or what Mike T likes to say 'A good time'"

  Case intRandom < 76
    strSelection =  "Panda Meat Cafe"

  Case intRandom < 78
    strSelection =  "Panera Bread"

  Case intRandom < 79
    strSelection =  "Roy Roger's"

  Case intRandom < 80
    strSelection =  "Wendy's"

  Case intRandom < 81
    strSelection =  "McDonald's"

  Case intRandom < 83
    strSelection =  "Canton Cafe"

  Case intRandom < 86
    strSelection =  "Uncle Charlie's"

  Case intRandom < 88
    strSelection =  "Cocorea Deli"

  Case intRandom < 90
    strSelection =  "Noodles and Company"

  Case intRandom < 92
    strSelection =  "Hard Times Cafe"

  Case intRandom <= 93
    strSelection =  "Dilly's Deli"

  Case intRandom <= 100
    strSelection = "Cafe del Rio"
End Select


wscript.echo Date() & " " & DotW() & " - " & strSelection





Function DotW()
Dim tmpDate

tmpDate =  DatePart("w",Date())

  Select Case tmpDate
    Case 2
    DotW = "Monday"
    Case 3
    DotW = "Tuesday"
    Case 4
    DotW = "Wednesday"
    Case 5
    DotW = "Thursday"
    Case 6
    DotW = "Friday"
  End Select

End Function
