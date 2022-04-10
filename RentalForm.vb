Option Explicit On
Option Strict On
Option Compare Binary
'Miranda Breves
'RCET0265
'Spring 2022
'Car Rental Program
'github url

Imports System.Text.RegularExpressions

Public Class RentalForm

    Function ValidateUserEntry(ByRef beginOdometer As Integer, ByRef endOdometer As Integer, ByRef day As Integer) As Boolean

        'Declaring variables
        Dim collectedErrorArray() As String = {""}

        Dim letterAndSpacePattern As String = "^[A-Za-z\s]+$"
        Dim lettersNumbersAndSpacePattern As String = "^[A-Za-z0-9\s]+$"
        Dim justNumbersPattern As String = "^[0-9]+$"
        Dim stringRegex As New Regex(letterAndSpacePattern)
        Dim stringNumberRegex As New Regex(lettersNumbersAndSpacePattern)
        Dim numberRegex As New Regex(justNumbersPattern)

        Dim zipCode As Integer

        Dim formattedErrorString As String = ""

        'name
        If NameTextBox.Text = "" Or Not stringRegex.IsMatch(NameTextBox.Text) Then
            collectedErrorArray(0) = "Customer Name"
            NameTextBox.Text = ""
        End If

        'address
        If AddressTextBox.Text = "" Or Not stringNumberRegex.IsMatch(AddressTextBox.Text) Then
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "Address"
            AddressTextBox.Text = ""
        End If

        'city
        If CityTextBox.Text = "" Or Not stringRegex.IsMatch(CityTextBox.Text) Then
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "City"
            CityTextBox.Text = ""
        End If

        'state
        If StateTextBox.Text = "" Or Not stringRegex.IsMatch(StateTextBox.Text) Then
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "State"
            StateTextBox.Text = ""
        End If

        'Zip code
        Try
            zipCode = CInt(ZipCodeTextBox.Text)

        Catch ex As Exception
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "Zip Code"
            ZipCodeTextBox.Text = ""
        End Try

        'odometers
        Try
            beginOdometer = CInt(BeginOdometerTextBox.Text)

            If beginOdometer < 0 Then
                If collectedErrorArray(0) <> "" Then
                    ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
                End If
                collectedErrorArray(collectedErrorArray.Length - 1) = "Beginning Odometer Reading"
                BeginOdometerTextBox.Text = ""
            End If

        Catch ex As Exception
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "Beginning Odometer Reading"
            BeginOdometerTextBox.Text = ""
        End Try

        Try
            endOdometer = CInt(EndOdometerTextBox.Text)

            If endOdometer < 0 Then

                If collectedErrorArray(0) <> "" Then
                    ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
                End If
                collectedErrorArray(collectedErrorArray.Length - 1) = "End Odometer Reading"
                EndOdometerTextBox.Text = ""

            End If

        Catch ex As Exception
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "End Odometer Reading"
            EndOdometerTextBox.Text = ""
        End Try

        'days
        Try
            day = CInt(DaysTextBox.Text)

            If day < 0 Then

                If collectedErrorArray(0) <> "" Then
                    ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
                End If
                collectedErrorArray(collectedErrorArray.Length - 1) = "Number of Days"
                DaysTextBox.Text = ""

            End If

        Catch ex As Exception
            If collectedErrorArray(0) <> "" Then
                ReDim Preserve collectedErrorArray(collectedErrorArray.Length)
            End If
            collectedErrorArray(collectedErrorArray.Length - 1) = "Number of Days"
            DaysTextBox.Text = ""
        End Try

        'Displaying any error messages that have been collected
        If collectedErrorArray(0) <> "" Then

            For i As Integer = 0 To collectedErrorArray.Length - 1
                If i <> 0 And i = collectedErrorArray.Length - 1 Then
                    formattedErrorString += $"and {collectedErrorArray(i)} "
                ElseIf collectedErrorArray.Length > 2 Then
                    formattedErrorString += $"{collectedErrorArray(i)}, "
                Else
                    formattedErrorString += $"{collectedErrorArray(i)} "
                End If
            Next

            If collectedErrorArray.Length > 0 Then
                MsgBox($"The {formattedErrorString}textboxes were empty or contained errors. Please review this " &
                   "information and correct the values.",, "Errors Found")
            Else
                MsgBox($"The {formattedErrorString}textbox was empty or contained errors. Please review this " &
                   "information and correct the value.",, "Error Found")
            End If

            Select Case collectedErrorArray(0)
                Case "Customer Name"
                    NameTextBox.Focus()
                Case "Address"
                    AddressTextBox.Focus()
                Case "City"
                    CityTextBox.Focus()
                Case "State"
                    StateTextBox.Focus()
                Case "Zip Code"
                    ZipCodeTextBox.Focus()
                Case "Beginning Odometer Reading"
                    BeginOdometerTextBox.Focus()
                Case "End Odometer Reading"
                    EndOdometerTextBox.Focus()
                Case "Number of Days"
                    DaysTextBox.Focus()
            End Select

            Return False
            Exit Function

        End If

        'Validating odometer readings
        If endOdometer < beginOdometer Then

            MsgBox("Please check the odometer readings; the beginning odometer reading must be smaller than the end " &
                   "odometer reading.",, "Odometer Reading Error")
            BeginOdometerTextBox.Text = ""
            EndOdometerTextBox.Text = ""
            BeginOdometerTextBox.Focus()

            Return False
            Exit Function

        End If

        'Validating days
        If day > 45 Then
            MsgBox("Please re-input your Number of Days value. The days cannot exceed 45.",, "Day Number Error")

            DaysTextBox.Text = ""
            DaysTextBox.Focus()

            Return False
            Exit Function
        End If

        Return True

    End Function

    Function CarCharges(ByVal customers As Integer, ByVal miles As Double, ByVal charges As Double) As String()

        Dim summaryArray(2) As String

        Static totalCustomers As Integer
        Static totalMiles As Double
        Static totalCharges As Double

        totalCustomers += customers
        totalMiles += miles
        totalCharges += charges

        summaryArray(0) = CStr(totalCustomers)
        summaryArray(1) = CStr(totalMiles)
        summaryArray(2) = CStr(totalCharges)

        Return summaryArray

    End Function

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MilesradioButton.Enabled = True

        'Figure out tool tips for all text boxes and buttons
        'RentalFormToolTip.SetToolTip()
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim beginOdometer, endOdometer, day, dayCharge As Integer
        Dim mileage, mileageCost, discount, totalCost As Double

        If ValidateUserEntry(beginOdometer, endOdometer, day) Then

            'mileage
            If MilesradioButton.Checked Then
                mileage = endOdometer - beginOdometer
            Else
                mileage = (endOdometer - beginOdometer) * 0.62  'converting from kilometers to miles
            End If

            TotalMilesTextBox.Text = CStr(mileage) & " mi"

            'mileage cost
            If mileage > 500 Then
                mileageCost = (mileage - 500) * 0.1
                mileageCost += 35.88 'If there is more than 500 miles, the 200-500 mile charge will be automatically be applied
            ElseIf mileage > 200 Then
                mileageCost = (mileage - 200) * 0.12
            Else
                mileageCost = 0
            End If

            MileageChargeTextBox.Text = Format(mileageCost, "c")

            'day charge
            dayCharge = day * 15
            DayChargeTextBox.Text = Format(dayCharge, "c")

            'discount
            If AAAcheckbox.Checked Then
                discount = 0.05
            End If

            If Seniorcheckbox.Checked Then
                discount += 0.03
            End If

            discount = (dayCharge + mileageCost) * discount
            TotalDiscountTextBox.Text = Format(discount, "c")

            'money owed
            totalCost = Math.Round((dayCharge + mileageCost - discount), 2)
            TotalChargeTextBox.Text = Format(totalCost, "c")

            CarCharges(1, mileage, totalCost)

        End If

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        'Clear input text boxes
        NameTextBox.Focus()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""

        'Set miles radio button back to default true & clear text boxes
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

        'Clear output text boxes
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""

    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim totalsArray(2) As String
        Dim displayString As String

        totalsArray = CarCharges(0, 0, 0)

        displayString = "Total Customers:" & vbTab & vbTab & totalsArray(0) & vbNewLine &
                        "Total Miles Driven:" & vbTab & vbTab & totalsArray(1) & " mi" & vbNewLine &
                        "Total Charges:" & vbTab & vbTab & "$" & totalsArray(2)

        MsgBox(displayString,, "Detailed Summary")

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

        Select Case MsgBox("Are you sure you want to exit?", vbYesNo, "Close")
            Case vbYes
                Me.Close()
            Case Else
                'If user doesn't want to exit, do nothing
        End Select

    End Sub

End Class
