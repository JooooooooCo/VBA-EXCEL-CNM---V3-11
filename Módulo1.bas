Attribute VB_Name = "Módulo1"
Sub SendRelease(what_address As String, subject_line As String, mail_body As String)

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")

    Dim olMail As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    
    olMail.SentOnBehalfOfName = "ITJ-PRONT-" & Plan1.Range("c9") & "@alianca.com.br"
'    olMail.To = "joao.costa@alianca.com.br"
    olMail.To = what_address
    olMail.subject = subject_line
    olMail.HTMLBody = mail_body
    olMail.CC = "ITJ-PRONT-" & Plan1.Range("C9")
    olMail.CC = olMail.CC + ";" + Plan1.Range("C11")
    olMail.Send
    

End Sub

Public Function procurashipper(shipper As String) As Boolean
    
    Dim row As Integer
    Dim achou As Boolean
    
    achou = False
    
    For row = 42 To 51
        If shipper = Plan1.Range("b" & row) Then
            achou = True
        End If
        
    Next row

    procurashipper = achou

End Function

Public Function procuraemail(email As String) As Boolean
    
    Dim row As Integer
    Dim achou As Boolean
    
    achou = False
    
    For row = 54 To 63
        If email = Plan1.Range("b" & row) Then
            achou = True
        End If

    Next row
    procuraemail = achou

End Function

Public Function countString(unidade As String) As Integer
    
    Dim unidades() As String
    
    unidades = Split(unidade, ",")
    
    countString = UBound(unidades) + 1
       

End Function

'---------------------------------------------- PRONTIDÃO DE CARGA 01 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 01 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 01 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 01 -----------------------------------------------------------------

Sub pront01()

Set arquivoexterno = Workbooks.Open(Plan1.Range("b6") & Plan1.Range("G9"))
Set Planilhaexterna = arquivoexterno.Sheets(1)

' limpando coluna email da linha O - agrm party

    Dim rowlimp As Integer
    
    For rowlimp = 13 To 2000
    
        If Planilhaexterna.Range("C" & rowlimp) = "O" Then
        
            Planilhaexterna.Range("D" & rowlimp).Select
            Selection.ClearContents
        
        End If
            
    Next rowlimp


' gravando email shp e forwarder
                  
    
    Dim rowemailshipper As Integer
    Dim rowemailforwarder As Integer
    Dim validaemail As String
        
        For rowemailshipper = 13 To 2000
        
              If Planilhaexterna.Range("C" & rowemailshipper) = "S" Then
                    
                    For rowemailforwarder = 13 To 2000
                    
                        If Planilhaexterna.Range("A" & rowemailforwarder) = Planilhaexterna.Range("A" & rowemailshipper) Then
                        
                                If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) = Planilhaexterna.Range("E" & rowemailshipper) Then
                                        
                                        validaemail = Planilhaexterna.Range("D" & rowemailforwarder)
                                        If procuraemail(validaemail) = False Then
                                        
                                            Planilhaexterna.Range("D" & rowemailshipper).Value = Planilhaexterna.Range("D" & rowemailshipper) & ";" & Planilhaexterna.Range("D" & rowemailforwarder)
                                        
                                        End If
                                End If
                                
                                If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) <> "" Then
                                
                                        If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) <> Planilhaexterna.Range("E" & rowemailshipper) Then
                                                
                                                Planilhaexterna.Range("E" & rowemailshipper).Value = Planilhaexterna.Range("E" & rowemailshipper) & " (" & Left(Planilhaexterna.Range("E" & rowemailforwarder), (Len(Planilhaexterna.Range("E" & rowemailforwarder))) - 3) & ")"
                                                validaemail = Planilhaexterna.Range("D" & rowemailforwarder)
                                                If procuraemail(validaemail) = False Then
                                                
                                                    Planilhaexterna.Range("D" & rowemailshipper).Value = Planilhaexterna.Range("D" & rowemailshipper) & ";" & Planilhaexterna.Range("D" & rowemailforwarder)
                                                
                                                End If
                                        End If
                                
                                End If
                        
                        End If
                    Next rowemailforwarder
        
            End If
        
        Next rowemailshipper


' macro

    Dim row As Integer
    Dim email As String
    Dim subject As String
    Dim mail_body_message As String
    Dim contbkg As String
    Dim contshp As String
    Dim shipper As String
    Dim rowshp As Integer
    
    Dim detalhebkg As String
            
    Dim qtdUnit As String
    Dim compQtdUnit As String
    

    shipper = ""
    email = ""
    contbkg = ""
    contshp = ""
    
For rowshp = 13 To 2000
    
If Planilhaexterna.Range("C" & rowshp) = "S" Then
  
    
    verificashp = contshp Like "*" & Planilhaexterna.Range("E" & rowshp) & "*"


     If verificashp = False Then
    

    
    
      For row = 13 To 2000
      
            If Planilhaexterna.Range("E" & row) = Planilhaexterna.Range("E" & rowshp) And Planilhaexterna.Range("C" & row) = "S" Then
            
                qtdUnit = countString(Planilhaexterna.Range("AA" & row))
                compQtdUnit = Planilhaexterna.Range("U" & row)
           

                detalhebkg = "<tr><td style='font-weight: bold;'>" & Planilhaexterna.Range("A" & row) & "</td><td>" & Planilhaexterna.Range("B" & row) & "</td><td>" & Planilhaexterna.Range("F" & row) & " " & Planilhaexterna.Range("H" & row) & " " & Planilhaexterna.Range("I" & row) & "</td><td>" & Planilhaexterna.Range("W" & row) & "</td><td>" & Planilhaexterna.Range("P" & row) & "</td><td>" & Planilhaexterna.Range("X" & row) & "</td><td>" & Planilhaexterna.Range("Y" & row) & "</td><td style='font-weight: bold; color: red;'>" & Planilhaexterna.Range("U" & row) & "</td><td>" & Planilhaexterna.Range("AA" & row) & "</td></tr>" & _
                                detalhebkg
                
                email = email & ";" & Planilhaexterna.Range("D" & row)
            
            End If
        
        
      Next row
        
                        
                tabeladetalhebkg = "<p><table border='1' style='width: 1180px; border-collapse: collapse; vertical-align: middle; text-align: center; font-size: 16px; font-family: Calibri, Candara, Segoe, Segoe UI, Optima, Arial, sans-serif;'><tr style='background-color: #003366; color: white'><td style='width: 100px;'>Booking</td><td style='width: 100px;'>Customer Ref.</td><td style='width: 120px;'>Vessel</td><td style='width: 160px;'>Port of Loading</td><td style='width: 100px;'>Estimated Sailing Date</td><td style='width: 160px;'>Port of Discharge</td><td style='width: 100px;'>Place of Delivery</td><td style='width: 110px;'>Total quantity of containers</td><td style='width: 230px;'>Container(s) No</td></tr>" & detalhebkg & "</table></p>"
                        
                        
                shipper = Planilhaexterna.Range("E" & rowshp)
                subject = Plan1.Range("B14")
                subject = Replace(subject, "substituirnavio", Planilhaexterna.Range("G13") & " " & Planilhaexterna.Range("H13") & Planilhaexterna.Range("I13"))
                subject = Replace(subject, "substituirporto", Planilhaexterna.Range("W13"))
                subject = Replace(subject, "substituirshipper", Planilhaexterna.Range("E" & rowshp))
                mail_body_message = Plan1.Range("B15")
                mail_body_message = Replace(mail_body_message, "substituirdetalhebkg", tabeladetalhebkg)
                mail_body_message = Replace(mail_body_message, "substituirtrade", Plan1.Range("C9"))
                mail_body_message = "<FONT FACE='Calibri'>" & mail_body_message & "</FONT>"
                
            If email <> "" Then
                If procurashipper(shipper) = False Then
                    
                        Call SendRelease(email, subject, mail_body_message)
                    
                End If
                
                detalhebkg = ""
                tabeladetalhebkg = ""
                email = ""
                
            End If
                        
        contshp = contshp & " - " & Planilhaexterna.Range("E" & rowshp)
     
     End If
End If

Next rowshp


    arquivoexterno.Close

End Sub
    

'---------------------------------------------- PRONTIDÃO DE CARGA 02 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 02 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 02 -----------------------------------------------------------------
'---------------------------------------------- PRONTIDÃO DE CARGA 02 -----------------------------------------------------------------

Sub pront02()

Set arquivoexterno = Workbooks.Open(Plan1.Range("b6") & Plan1.Range("G9"))
Set Planilhaexterna = arquivoexterno.Sheets(1)

' limpando coluna email da linha O - agrm party

    Dim rowlimp As Integer
    
    For rowlimp = 13 To 2000
    
        If Planilhaexterna.Range("C" & rowlimp) = "O" Then
        
            Planilhaexterna.Range("D" & rowlimp).Select
            Selection.ClearContents
        
        End If
            
    Next rowlimp


' gravando email shp e forwarder
                  
    
    Dim rowemailshipper As Integer
    Dim rowemailforwarder As Integer
    Dim validaemail As String
        
        For rowemailshipper = 13 To 2000
        
              If Planilhaexterna.Range("C" & rowemailshipper) = "S" Then
                    
                    For rowemailforwarder = 13 To 2000
                    
                        If Planilhaexterna.Range("A" & rowemailforwarder) = Planilhaexterna.Range("A" & rowemailshipper) Then
                        
                                If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) = Planilhaexterna.Range("E" & rowemailshipper) Then
                                        
                                        validaemail = Planilhaexterna.Range("D" & rowemailforwarder)
                                        If procuraemail(validaemail) = False Then
                                        
                                            Planilhaexterna.Range("D" & rowemailshipper).Value = Planilhaexterna.Range("D" & rowemailshipper) & ";" & Planilhaexterna.Range("D" & rowemailforwarder)
                                        
                                        End If
                                End If
                                
                                If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) <> "" Then
                                
                                        If Planilhaexterna.Range("C" & rowemailforwarder) = "F" And Planilhaexterna.Range("E" & rowemailforwarder) <> Planilhaexterna.Range("E" & rowemailshipper) Then
                                                
                                                Planilhaexterna.Range("E" & rowemailshipper).Value = Planilhaexterna.Range("E" & rowemailshipper) & " (" & Left(Planilhaexterna.Range("E" & rowemailforwarder), (Len(Planilhaexterna.Range("E" & rowemailforwarder))) - 3) & ")"
                                                validaemail = Planilhaexterna.Range("D" & rowemailforwarder)
                                                If procuraemail(validaemail) = False Then
                                                
                                                    Planilhaexterna.Range("D" & rowemailshipper).Value = Planilhaexterna.Range("D" & rowemailshipper) & ";" & Planilhaexterna.Range("D" & rowemailforwarder)
                                                
                                                End If
                                        End If
                                
                                End If
                        
                        End If
                    Next rowemailforwarder
        
            End If
        
        Next rowemailshipper

' macro


    Dim row As Integer
    Dim email As String
    Dim subject As String
    Dim mail_body_message As String
    Dim contbkg As String
    Dim contshp As String
    Dim shipper As String
    Dim rowshp As Integer
    Dim rowemail As Integer
      
    Dim detalhebkg As String
            
    Dim qtdUnit As String
    Dim compQtdUnit As String
    

    shipper = ""
    email = ""
    contbkg = ""
    contshp = ""
    
For rowshp = 13 To 2000
    
If Planilhaexterna.Range("C" & rowshp) = "S" Then
  
    
    verificashp = contshp Like "*" & Planilhaexterna.Range("E" & rowshp) & "*"


     If verificashp = False Then
    

    
    
      For row = 13 To 2000
      
            If Planilhaexterna.Range("E" & row) = Planilhaexterna.Range("E" & rowshp) And Planilhaexterna.Range("C" & row) = "S" Then
            
                qtdUnit = countString(Planilhaexterna.Range("AA" & row))
                compQtdUnit = Planilhaexterna.Range("U" & row)
                
                If qtdUnit <> compQtdUnit Then
                
                    

                    detalhebkg = "<tr><td style='font-weight: bold;'>" & Planilhaexterna.Range("A" & row) & "</td><td>" & Planilhaexterna.Range("F" & row) & " " & Planilhaexterna.Range("H" & row) & " " & Planilhaexterna.Range("I" & row) & "</td><td>" & Planilhaexterna.Range("W" & row) & "</td><td>" & Planilhaexterna.Range("P" & row) & "</td><td>" & Planilhaexterna.Range("X" & row) & "</td><td>" & Planilhaexterna.Range("Y" & row) & "</td><td>" & compQtdUnit & "</td><td style='font-weight: bold; color: red;'>" & qtdUnit & "</td><td>" & Planilhaexterna.Range("AA" & row) & "</td></tr>" & _
                                    detalhebkg
                
                    
                    email = email & ";" & Planilhaexterna.Range("D" & row)
                                    
                End If
                
            End If
        
        
      Next row
                
                
                tabeladetalhebkg = "<p><table border='1' style='width: 1180px; border-collapse: collapse; vertical-align: middle; text-align: center; font-size: 16px; font-family: Calibri, Candara, Segoe, Segoe UI, Optima, Arial, sans-serif;'><tr style='background-color: #003366; color: white'><td style='width: 100px;'>Booking</td><td style='width: 120px;'>Vessel</td><td style='width: 160px;'>Port of Loading</td><td style='width: 100px;'>Estimated Sailing Date</td><td style='width: 160px;'>Port of Discharge</td><td style='width: 100px;'>Place of Delivery</td><td style='width: 105px;'>Total of Booked Container(s)</td><td style='width: 105px;'>Total of Picked Up Container(s)</td><td style='width: 230px;'>Container(s) No</td></tr>" & detalhebkg & "</table></p>"
                        
                        
                shipper = Planilhaexterna.Range("E" & rowshp)
                subject = Plan1.Range("H14")
                subject = Replace(subject, "substituirnavio", Planilhaexterna.Range("G13") & " " & Planilhaexterna.Range("H13") & Planilhaexterna.Range("I13"))
                subject = Replace(subject, "substituirporto", Planilhaexterna.Range("W13"))
                subject = Replace(subject, "substituirshipper", Planilhaexterna.Range("E" & rowshp))
                mail_body_message = Plan1.Range("H15")
                mail_body_message = Replace(mail_body_message, "substituirdetalhebkg", tabeladetalhebkg)
                mail_body_message = Replace(mail_body_message, "substituirnavio", Planilhaexterna.Range("G13") & " " & Planilhaexterna.Range("H13") & Planilhaexterna.Range("I13"))
                mail_body_message = Replace(mail_body_message, "substituirporto", Planilhaexterna.Range("W13"))
                mail_body_message = Replace(mail_body_message, "substituirddl", Plan1.Range("J9"))
                mail_body_message = Replace(mail_body_message, "substituirtrade", Plan1.Range("C9"))
                mail_body_message = "<FONT FACE='Calibri'>" & mail_body_message & "</FONT>"
            
            If email <> "" Then
                
                If procurashipper(shipper) = False Then
                        Call SendRelease(email, subject, mail_body_message)
                End If
                
                detalhebkg = ""
                tabeladetalhebkg = ""
                email = ""
               
            End If
                        
        contshp = contshp & " - " & Planilhaexterna.Range("E" & rowshp)
        
     
     End If
    
End If

Next rowshp


    arquivoexterno.Close

End Sub
