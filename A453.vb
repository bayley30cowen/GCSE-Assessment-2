Public Class Form1
'This declares all variables publicly so all subs can see these variables. 
Dim Date1, Date2 As Date
Dim timediff, averageSpeed As Double
Dim roadName, firstName, surname, homeAddress, standardCheck, regnum, title As String
Dim roadLength, speedLimit As Integer
    
Sub Dimming()
'This declares all variables in one sub and can be enabled in all subs. This is used for efficiency    
  roadName = Me.TextBox3.Text
  speedLimit = Me.TextBox1.Text
  regnum = Me.TextBox2.Text
  Date1 = Me.MaskedTextBox1.Text
  Date2 = Me.MaskedTextBox2.Text
 roadLength = Me.TextBox4.Text
End Sub
Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
'This allows all dims from the sub dimming to be present in this sub
        Dimming()
  'This uses the function to collect data of the driver from a CSV     
 If functionread(Me.TextBox2.Text) Then
'This tells the form the destination the data will be collected from            
Using MyReader As New 
Microsoft.VisualBasic.FileIO.TextFieldParser("s:\names.csv")
MyReader.TextFieldType = FileIO.FieldType.Delimited
'This sets the delimeter in the CSV as a comma               
 MyReader.SetDelimiters(",")
                Dim currentRowSelected As String()
                While Not MyReader.EndOfData
                    currentRowSelected = MyReader.ReadFields()
'This tells the form if the registration plate on the form matches a registration plate in the CSV all the details with the plate will be outputted.
                    If currentRowSelected(0) = regnum Then
'This creates the message the user will see about the driver’s information
MsgBox(currentRowSelected(4) & " " & currentRowSelected(2) &  " " & currentRowSelected(3) & " " & "lives at" & " " & currentRowSelected(1))
                        homeAddress = currentRowSelected(1)
                        firstName = currentRowSelected(2)
                        surname = currentRowSelected(3)
                        title = currentRowSelected(4) 
'This runs the sub routine that sends the data to a CSV.
               If speedLimit < averageSpeed Then
                storedetailsnonstandard()
                MessageBox.Show(firstName & " " & surname & " " & "has been sent a fine", "Fine Alert")
            Else : MessageBox.Show("The driver cannot be sent a fine as he is not breaking the limit", "Fine Alert")
                    End If
                End While
            End Using
        End If
'If the plate is non-standard it will try to find it without using the mask set in the function
Else : Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("c:\test\names.csv")
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRowSelected As String()
                While Not MyReader.EndOfData
                    currentRowSelected = MyReader.ReadFields()
                    If currentRowSelected(0) = regnum Then
                        MessageBox.Show(currentRowSelected(4) & " " & currentRowSelected(2) & " " & currentRowSelected(3) & " " & "lives at" & " " & currentRowSelected(1), "Drivers Info")
                        homeAddress = currentRowSelected(1)
                        firstName = currentRowSelected(2)
                        surname = currentRowSelected(3)
                        title = currentRowSelected(4)
                    End If
End If
End While
            End Using
            If speedLimit < averageSpeed Then
                storedetailsnonstandard()
                MessageBox.Show(firstName & " " & surname & " " & "has been sent a fine", "Fine Alert")
            Else : MessageBox.Show("The driver cannot be sent a fine as he is not breaking the limit", "Fine Alert")
            End If    
End Sub


 'This is a sub routine that sends the data on the form and the info from the CSV to a CSV if the plate is standard
Sub storedetailsstandard()
        Dim writeToCSV As System.IO.StreamWriter
        writeToCSV = My.Computer.FileSystem.OpenTextFileWriter("s:\driverinfostandard.csv", True)
        writeToCSV.WriteLine(Me.TextBox3.Text & "," & homeAddress & "," & firstName & "," & surname & "," & averageSpeed & "mph" & "," & averageSpeed - speedLimit & "mph over the limit")
    End Sub

 'This is a sub routine that sends the data on the form and the info from the CSV to a CSV if the plate is non-standard
Sub storedetailsnonstandard()
        Dim writeToCSV As System.IO.StreamWriter
        writeToCSV = My.Computer.FileSystem.OpenTextFileWriter("s:\driverinfononstandard.csv", True)
        writeToCSV.WriteLine(Me.TextBox3.Text & "," & homeAddress & "," & firstName & "," & surname & "," & averageSpeed & "mph" & "," & averageSpeed - speedLimit & "mph over the limit")
    End Sub


Function functionread(ByVal regnum As String) As Boolean
'This creates the function that is used to find the registration plate in the CSV with the driver’s information. This checks that each part of the registration plate, the letters and numbers match.
        Dimming()
        Dim upper As New System.Text.RegularExpressions.Regex("[A-Z]")
        Dim number As New System.Text.RegularExpressions.Regex("[0-9]")
'This states if the registration plates length does not equal 7 then do not continue
        If Len(regnum) <> 7 Then Return False
'This checks if the registration on the form matches a registration plate on the CSV.
        If upper.Matches(Mid(regnum, 1, 2)).Count < 2 Then Return False
        If number.Matches(Mid(regnum, 3, 2)).Count < 2 Then Return False
        If upper.Matches(Mid(regnum, 5, 3)).Count < 3 Then Return False
        Return True
    End Function




Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
Dimming()              
 'This if statement validates if any fields are empty and breaks the code if there is an empty field
        If Me.TextBox3.Text.Length = 0 Then
            'This creates the alert to alert the user that a field is empty
MsgBox("Please enter the name of the road", vbInformation, "Validation Alert")
'By leaving the sub it will stop further code, so it will not break later on            
Exit Sub
        ElseIf Me.TextBox1.Text.Length = 0 Then
           
 MsgBox("Please enter the speed limit", vbInformation, "Validation Alert")
            Exit Sub
        ElseIf Me.TextBox4.Text.Length = 0 Then
            MsgBox("Please enter the distance between sites", vbInformation, "Validation Alert")
            Exit Sub
        ElseIf Me.TextBox2.Text.Length = 0 Then
            MsgBox("Please enter the registration plate", vbInformation, "Validation Alert")
            Exit Sub
        ElseIf Not MaskedTextBox1.MaskCompleted Then
            MsgBox("Please enter the time the vehicle passed site one", vbInformation, "Validation Alert")
            Exit Sub
        ElseIf Not MaskedTextBox2.MaskCompleted Then
            MsgBox("Please enter the time the vehicle passed site two", vbInformation, "Validation Alert")
            Exit Sub
        End If

        'This if statement validates that the correct values are entered into the textbox's e.g numerical values for speed limit
        If Not IsNumeric(Me.TextBox1.Text) Then
            MsgBox("Invalid data entered, Please enter numbers only for the speed limit.", vbInformation, "Alert")
            Exit Sub
        ElseIf IsNumeric(Me.TextBox3.Text) Then
            MsgBox("Invalid data entered, Please enter letters only for the road name.", vbInformation, "Alert")
            Exit Sub
        ElseIf Not IsNumeric(Me.TextBox4.Text) Then
            MsgBox("Invalid data entered, Please enter numbers only for the distance between sites.", vbInformation, "Alert")
            Exit Sub
        End If
       


        
        'this toggles the driver's info button to be enabled
        Button1.Enabled = True
        'This creates the time difference between date1 and date2
        timediff = DateDiff("s", Date1, Date2)
        'This calculates the average speed of the vehicle I used one calculation for efficiency.
        averageSpeed = roadLength * 3600 / timediff
        'This will output the speed of the vehicle efficiently
        MessageBox.Show("The speed of the vehicle is " & averageSpeed & "mph", "The speed of the vehicle")
        'This shows the user the mph the vehicle was over the limit
        If averageSpeed > speedLimit Then
            MessageBox.Show("The vehicle is " & averageSpeed - speedLimit & "mph over the limit!", "Breaking Speed Alert")
        Else
            MessageBox.Show("The vehicle is not breaking the speed limit", "Alert")
        End If
        'This sends data to the listboxs so the user can look at all necessary data
        If speedLimit < averageSpeed Then
            Me.ListBox1.Items.Add(regnum)
            Me.ListBox2.Items.Add(averageSpeed)
            Me.ListBox3.Items.Add(averageSpeed - speedLimit & "mph over the limit")
        End If
        

'This checks whether the registration number fits the standard mask from 2001- and the original mask used from 1983 - 2001
        
If speedLimit < averageSpeed Then
            If UCase$(regnum) Like "[A-Z][A-Z][0-9][0-9][A-Z][A-Z][A-Z]" Or UCase$(regnum) Like "[A-Z][0-9][0-9][0-9][A-Z][A-Z][A-Z]" Then
                Me.ListBox4.Items.Add("Standard Plate")
                standardCheck = "standard"
            Else
                Me.ListBox4.Items.Add("Non-Standard Plate")
                standardCheck = "nonstandard"
            End If
        End If
    End Sub
  Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'This makes the registration number textbox all uppercase letters
        Me.TextBox2.CharacterCasing = CharacterCasing.Upper
    End Sub




Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'This creates the reset button that will reset every textbox and listbox.
        Controls.Clear()
        InitializeComponent()
  End Sub

Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
'This creates the close button       
 Close()
    End Sub
End Class
