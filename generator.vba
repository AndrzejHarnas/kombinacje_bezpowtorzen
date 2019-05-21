Sub genkomb()

n = 4
m = 2

z = Silnia(n)
x = Ikombezpow(z, n, m)

MsgBox "ilość kombinacji: " & x

v = GLKBP(n, m, x)



End Sub
'Funkcja obliczająca silnie
Function Silnia(wielkosczbioru)

If wielkosczbioru = 1 Or wielkosczbioru = 0 Then

Silnia = 1

Else

Silnia = wielkosczbioru * Silnia(wielkosczbioru - 1)

End If

End Function

'ilosc kombinacji bez powtórzen

Function Ikombezpow(silnian, wielkosczbioru, wielkoscpodzbioru)

'wzór: n! /(n-m)!*m!
' n=wielkosc zbioru m = wielkosc pozdzbioru

mianownik = Silnia(wielkosczbioru - wielkoscpodzbioru) * Silnia(wielkoscpodzbioru)
licznik = silnian

Ikombezpow = licznik / mianownik

End Function

'generowanie k losowych kombinacji bez poqwtórzeń

Function GLKBP(wielkosczbioru, wielkoscpodzbioru, k)

Dim tmp(20, 1000) 'pamieć już wylosowanych kombinacji
Dim podzbior(20) 'podzbior przechowujący wygenerowaną losowo kombinacje bez powtorzeń
Dim i, j, z, x, e As Integer 'liczniki pętli
Dim zmp As Integer 'wylosowana pojedyncza liczba w zakresie od 1 do wielkosczbioru
Dim powtorka As Integer 'zmienna odpowiedzialna za losowanie podzbioru, tak długo az to będzie właściwa z punktu widzenia matematycznego kombinacja wielkoscpodzbioru elementowa bez powtórzeń ze zbioru wielkosczbioru elementowego
Dim powtorka2 As Integer 'powtarzanie generowania kombinacji bez powtórzeń, tak długo aż bedzie się powtarzać. Celem zbioru jest uzyskanie nie wygenerowanej wcześniej kombinacji bez powtórzeń
Dim ilosc As Integer 'zmienna odpowiedzialna za wylosowanie tylko wielkoscpodzbioru elementów do podzbioru podzbior
Dim przecinek As Integer 'czy wypisywac przecinek
Dim wynik As String


'zerowanie pamieci zapamietanych kombinacji

For i = 1 To 1000

    For z = 1 To wielkoscpodzbioru
    
    tmp(z - 1, i) = 0
    
    Next z

Next i

'zerowanie podzbioru

For i = 1 To wielkoscpodzbioru

podzbior(i - 1) = 0

Next i

For j = 1 To k

    przecinek = 0
    ilosc = 0
    
    'losowanie kombinacji wielkoscpodzbioru elementowej bez powtórzeń ze zbioru wielkosczbioru elementowego
    powtorka2 = 0
        
    Do While powtorka2 = 0
        
            'wylosowanie prawidłowej z punktu widzenia matematyki kombinacji bez powtórzeń
            Do While ilosc < wielkoscpodzbioru
            
                powtorka = 0
                Randomize
                zmp = 1 + Round(Rnd * (wielkosczbioru - 1), 1)
                
                For z = 1 To wielkoscpodzbioru
                
                    If podzbior(z - 1) = zmp Then
                    
                    powtorka = 1
                    
                    Exit For
                    
                    End If
                
                Next z
                
                If powtorka = 0 Then
                
                    ilosc = ilosc + 1
                    podzbior(ilosc - 1) = zmp
                
                End If
        
            Loop
                
        'sprawdzenie czy ta kombinacja bez powtorzen była już wczesniej wylosowana i jak tak to powtórka losowania
        powtorka2 = 1
        
        For i = 1 To k
        
            If tmp(0, i - 1) = 0 Then
            
            Exit Do
            
            End If
            
            For z = 1 To wielkoscpodzbioru
            
                For x = 1 To wielkoscpodzbioru
                
                    If tmp(z - 1, i - 1) = podzbior(x - 1) Then
                        powtorka2 = 0
                        Exit For
                    End If
                    
                    powtorka2 = 1
                
                Next x
                
                If powtorka2 = 1 Then
                
                Exit For
                
                End If
            
            Next z
         
            If powtorka2 = 0 Then
            
                For e = 1 To wielkoscpodzbioru
                
                    podzbior(e - 1) = 0
                
                Next e
                ilosc = 0
                
                Exit For
            End If
            
           
         
         Next i
        
        
   Loop
        
   'wypisywanie i zapamietanie kombinacji bez powtorzen
   For i = 1 To wielkoscpodzbioru
        
        wynik = "" & wynik & " " & podzbior(i - 1) & ""
        tmp(i - 1, j - 1) = podzbior(i - 1)
        przecinek = przecinek + 1
        
        If przecinek = wielkoscpodzbioru Then
            MsgBox wynik
            wynik = ""
        End If
        
   Next i
   
   'ponowne zerowanie podzbioru i przygotowanie do następnego generowania
   
   For i = 1 To wielkoscpodzbioru
   
   podzbior(i - 1) = 0
   
   Next i
   
Next j

End Function