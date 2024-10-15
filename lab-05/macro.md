---
marp: true
theme: default
class: invert
---

# Wstęp

Makro, które zaprezentujemy tworzy nowy arkusz, **tworzy tabelę** o wymiarach podanych przez użytkownika oraz pozwala po kolei **wypełnić dane komórki** wartościami podanymi przez użytkownika.

---

## Całe makro

```vb
1	Sub CreateTableAndFill()
2	    Dim numRows As Integer
3	    Dim numCols As Integer
4	    Dim currentRow As Integer
5	    Dim currentCol As Integer
6	    Dim userInput As Variant
7    
8	    numRows = InputBox("Enter the number of rows:", "Table Dimensions")
9	    numCols = InputBox("Enter the number of columns:", "Table Dimensions")
10
11	    If IsNumeric(numRows) And IsNumeric(numCols) And numRows > 0 And numCols > 0 Then
12	        Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Table"
13	        Set ws = ActiveSheet
14        
15	        For currentRow = 1 To numRows
16	            For currentCol = 1 To numCols
17	                userInput = InputBox("Enter value for cell [" & currentRow & ", " & currentCol & "]:", "Enter Value")
18                
19	                If userInput <> "" Then
20	                    ws.Cells(currentRow, currentCol).Value = userInput
21	                Else
22	                    Exit Sub
23	                End If
24	            Next currentCol
25	        Next currentRow
26	    Else
27	        MsgBox "Please enter valid numeric values for rows and columns.", vbExclamation
28	    End If
29	End Sub
```

---

## Deklaracja zmiennych
* W linijkach ```2```-```6``` deklarujemy zmienne 
	* ```numRows``` - będzie zawierać ilość wierszy
	* ```numCols``` - będzie zawierać ilość kolumn
	* ```currentRow``` - będzie zawierać numer obecnego wiersza
	* ```currentCol``` - będzie zawierać numer obecnej kolumny
	* ```userInput``` - będzie zawierać dowolny input od użytkownika


```vb
2	Dim numRows As Integer
3	Dim numCols As Integer
4	Dim currentRow As Integer
5	Dim currentCol As Integer
6	Dim userInput As Variant
```

---

## Podanie wartości
* W linijkach ```9``` i ```10``` zapisujemy do zmiennych ```numRows``` i ```numCols``` ilość wierszy i kolumn podanych przez użytkownika
* Wykorzystamy to do stworzenia tabeli o podanych wymiarach

```vb
9	numRows = InputBox("Enter the number of rows:", "Table Dimensions")
10	numCols = InputBox("Enter the number of columns:", "Table Dimensions")
```

---

## Weryfikacja danych
* W linijce ```11``` sprawdzamy, czy podane przez użytkownika wartości są wartościami numerycznymi oraz czy są większe niż ```0```

```vb
11	If IsNumeric(numRows) And IsNumeric(numCols) Then
		'[...] - następne slajdy
26	Else
27		MsgBox "Please enter valid numeric values for rows and columns.", vbExclamation
28	End If
```

---

## Stworzenie arkusza z tabelką
* W linijkach ```12``` i ```13``` tworzymy nowy arkusz o nazwie ```"Table"``` i ustawiamy go jako aktywny arkusz

```vb
12	Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Table"
13	Set ws = ActiveSheet
```

---

## Wpisywanie wartości cz. 1
* Tworzymy dwie pętle ```for``` (jedna zagnieżdżona w drugiej)
* Pozwolą nam one na przechodzenie po kolejnych wierszach, w kolejnych kolumnach

```vb
15	For currentRow = 1 To numRows
16		For currentCol = 1 To numCols
			'[...] - następny slajd
24		Next currentCol
25	Next currentRow
```

---

## Wpisywanie wartości cz. 2
* W linijce ```17``` prosimy użytkownika o podanie wartości dla danej komórki
* Następnie w linijce ```19``` sprawdzamy czy użytkownik wpisał wartość, czy kliknął przycisk ```Cancel```
	* Jeżeli wpisał wartość, to realizujemy linijkę ```20``` - przypisujemy danej komórce wartość podaną przez użytkownika
	* Jeżeli kliknął ```Cancel```, to kończymy makro

```vb
17	userInput = InputBox("Enter value for cell [" & currentRow & ", " & currentCol & "]:", "Enter Value")
18                
19	If userInput <> "" Then
20		ws.Cells(currentRow, currentCol).Value = userInput
21	Else
22		Exit Sub
23	End If
```

---

# Koniec

**Autorstwa:**
- Zachariasz Jażdżewski
- Bartosz Guzowski
- Justyna Hinz
- Wiktoria Grabara