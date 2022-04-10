'Miranda Breves
'RCET0265
'Spring 2022
'Car Rental Program
'https://github.com/MEBreves/CarRental

Option Explicit On
Option Strict On
Option Compare Binary

Imports System.Text.RegularExpressions

Public Class RentalForm

    'This function collects error strings and adds them to an array.
    Sub UpdateErrorArray(ByRef errorArray() As String, ByVal errorString As String)

        'In visual basic, an array has to be declared as a set size. Redimensioning an array allows new
        'values to be added to it while preserving it keeps the values it already has.
        If errorArray(0) <> "" Then
            ReDim Preserve errorArray(errorArray.Length)
        End If
        errorArray(errorArray.Length - 1) = errorString

    End Sub

    'This function is used to check all of the customer and rental car data entered. If any values are
    'flagged as empty or incorrect, an error message will be displayed to the user and all erroneous
    'values will be cleared.
    Function ValidateUserEntry(ByRef beginOdometer As Integer, ByRef endOdometer As Integer, ByRef day As Integer) As Boolean

        'Declaring variables
        Dim collectedErrorArray() As String = {""}
        Dim zipCode As Integer
        Dim formattedErrorString As String = ""

        'A regex (regular expression) is used to verify the customer's data. It will return a true if
        'the user's values match the pattern set to the regex. The patterns set below are for letters and
        'spaces, and letters, numbers, and spaces, respectively. Any other characters aside from those
        'in the user's values will cause the regex to return false and raise an error.
        Dim letterAndSpacePattern As String = "^[A-Za-z\s]+$"
        Dim lettersNumbersAndSpacePattern As String = "^[A-Za-z0-9\s]+$"
        Dim stringRegex As New Regex(letterAndSpacePattern)
        Dim stringNumberRegex As New Regex(lettersNumbersAndSpacePattern)


        'The customer's name will be verified with a letter and space pattern regex. Errors will be
        'collected and the user's input will be cleared.
        If NameTextBox.Text = "" Or Not stringRegex.IsMatch(NameTextBox.Text) Then
            UpdateErrorArray(collectedErrorArray, "Customer Name")
            NameTextBox.Text = ""
        End If

        'The address will be verified with a letter, number, and string regex. Errors will be
        'collected and the user's input will be cleared if so.
        If AddressTextBox.Text = "" Or Not stringNumberRegex.IsMatch(AddressTextBox.Text) Then
            UpdateErrorArray(collectedErrorArray, "Address")
            AddressTextBox.Text = ""
        End If

        'The city and state will be verified with a letter and space regex. Errors will be
        'collected and the user's input will be cleared if so.
        If CityTextBox.Text = "" Or Not stringRegex.IsMatch(CityTextBox.Text) Then
            UpdateErrorArray(collectedErrorArray, "City")
            CityTextBox.Text = ""
        End If

        If StateTextBox.Text = "" Or Not stringRegex.IsMatch(StateTextBox.Text) Then
            UpdateErrorArray(collectedErrorArray, "State")
            StateTextBox.Text = ""
        End If

        'The Zip code, odometer readings, and days will be verified as numbers greater than zero with a
        'try-catch. Errors will be collected and the user's input will be cleared if so.
        Try     'Checking the Zip Code
            zipCode = CInt(ZipCodeTextBox.Text)

            If zipCode < 0 Then     'If the zip code is negative, an error will also be recorded
                UpdateErrorArray(collectedErrorArray, "Zip Code")
                ZipCodeTextBox.Text = ""
            End If

        Catch ex As Exception
            UpdateErrorArray(collectedErrorArray, "Zip Code")
            ZipCodeTextBox.Text = ""
        End Try

        Try     'Checking Beginning Odometer reading
            beginOdometer = CInt(BeginOdometerTextBox.Text)

            If beginOdometer < 0 Then   'If the reading is negative, an error will also be recorded
                UpdateErrorArray(collectedErrorArray, "Beginning Odometer Reading")
                BeginOdometerTextBox.Text = ""
            End If

        Catch ex As Exception
            UpdateErrorArray(collectedErrorArray, "Beginning Odometer Reading")
            BeginOdometerTextBox.Text = ""
        End Try

        Try     'Checking Ending Odometer reading
            endOdometer = CInt(EndOdometerTextBox.Text)

            If endOdometer < 0 Then     'If the reading is negative, an error will also be recorded
                UpdateErrorArray(collectedErrorArray, "End Odometer Reading")
                EndOdometerTextBox.Text = ""
            End If
        Catch ex As Exception
            UpdateErrorArray(collectedErrorArray, "End Odometer Reading")
            EndOdometerTextBox.Text = ""
        End Try

        Try     'Checking the number of days
            day = CInt(DaysTextBox.Text)

            If day < 1 Then 'If the number of days is zero or less, an error will be recorded
                UpdateErrorArray(collectedErrorArray, "Number of Days")
                DaysTextBox.Text = ""
            End If

        Catch ex As Exception
            UpdateErrorArray(collectedErrorArray, "Number of Days")
            DaysTextBox.Text = ""
        End Try

        'Displaying any error messages that have been collected
        If collectedErrorArray(0) <> "" Then

            'The collected errors will be formatted into a string based on how many errors there are
            For i As Integer = 0 To collectedErrorArray.Length - 1
                If i <> 0 And i = collectedErrorArray.Length - 1 Then
                    formattedErrorString += $"and {collectedErrorArray(i)} "
                ElseIf collectedErrorArray.Length > 2 Then
                    formattedErrorString += $"{collectedErrorArray(i)}, "
                Else
                    formattedErrorString += $"{collectedErrorArray(i)} "
                End If
            Next

            'A message box will display the errors based on how many errors there are
            If collectedErrorArray.Length > 0 Then
                MsgBox($"The {formattedErrorString}textboxes were empty or contained errors. Please review this " &
                   "information and correct the values.",, "Errors Found")
            Else
                MsgBox($"The {formattedErrorString}textbox was empty or contained errors. Please review this " &
                   "information and correct the value.",, "Error Found")
            End If

            'The first textbox with an error will be focused on so the user can begin to resolve values
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

            'As errors were found in the user's values, the function will return false as values weren't
            'validated
            Return False
            Exit Function

        End If

        'The odometer readings will need to be checked to see if the end is greater than the beginning.
        If endOdometer < beginOdometer Then

            MsgBox("Please check the odometer readings; the beginning odometer reading must be smaller than the end " &
                   "odometer reading.",, "Odometer Reading Error")
            BeginOdometerTextBox.Text = ""
            EndOdometerTextBox.Text = ""
            BeginOdometerTextBox.Focus()

            'As there was an issue with the odometer readings, the user's values aren't valid and the
            'function will return false
            Return False
            Exit Function

        End If

        'The number of days will need to be less than 45, or else the program will not run.
        If day > 45 Then
            MsgBox("Please re-input your Number of Days value. The days cannot exceed 45.",, "Day Number Error")

            DaysTextBox.Text = ""
            DaysTextBox.Focus()

            'If the days were greater than 45, the user's values are not valid and the function returns false.
            Return False
            Exit Function
        End If

        'If no errors were found and all values were validated, then the validation succeeded and
        'the function will return true
        Return True

    End Function

    'This function is used to retain the scope of the total customer, miles, and charges variables so that
    'global variables are not used. Subs can store variables by placing appropriate values in the parameters,
    'or retrive values through the returned string array.
    Function CarCharges(ByVal customers As Integer, ByVal miles As Double, ByVal charges As Double) As String()

        'The array to be returned containing the total customers, miles, and charges values. The array to
        'return must contain strings instead of numbers because the variables are different number types - 
        'integer and double.
        Dim summaryArray(2) As String

        'Declaring the variables as static to retain values while the program is running
        Static totalCustomers As Integer
        Static totalMiles, totalCharges As Double

        'Updating the variables based off of the input parameters
        totalCustomers += customers
        totalMiles += miles
        totalCharges += charges

        'Placing the updated variables into the array to be returned
        summaryArray(0) = CStr(totalCustomers)
        summaryArray(1) = CStr(totalMiles)
        summaryArray(2) = CStr(totalCharges)

        Return summaryArray

    End Function

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'When the form loads, the miles radio button will be set checked by default
        MilesradioButton.Checked = True

        'The tooltips for all textboxes and buttons are created and applied on load.
        RentalFormToolTip.SetToolTip(NameTextBox, "The full name, first and last, of the customer.")
        RentalFormToolTip.SetToolTip(AddressTextBox, "The customer's billing address.")
        RentalFormToolTip.SetToolTip(CityTextBox, "The customer's city to be addressed to.")
        RentalFormToolTip.SetToolTip(StateTextBox, "The customer's state to be addressed to.")
        RentalFormToolTip.SetToolTip(ZipCodeTextBox, "The customer's zip code to be addressed to.")
        RentalFormToolTip.SetToolTip(BeginOdometerTextBox, "The beginning odometer reading of the rental car.")
        RentalFormToolTip.SetToolTip(EndOdometerTextBox, "The ending odometer reading of the rental car " &
                                                         "after the customer has returned it.")
        RentalFormToolTip.SetToolTip(DaysTextBox, "The number of days the customer used the rental car.")


        RentalFormToolTip.SetToolTip(TotalMilesTextBox, "The total number of miles the customer drove the car.")
        RentalFormToolTip.SetToolTip(MileageChargeTextBox, "The cost of the miles driven by the customer, in USD.")
        RentalFormToolTip.SetToolTip(DayChargeTextBox, "The cost for the customer to use the car for all the" &
                                                        " days it was gone.")
        RentalFormToolTip.SetToolTip(TotalDiscountTextBox, "The amount the AAA and/or the senior discount " &
                                                           "will reduce the total price by.")
        RentalFormToolTip.SetToolTip(TotalChargeTextBox, "The total charge the customer owes for renting the" &
                                                         "car.")


        RentalFormToolTip.SetToolTip(MilesradioButton, "To be checked if the car's odometer is in miles.")
        RentalFormToolTip.SetToolTip(KilometersradioButton, "To be checked if the car's odometer is in " &
                                                            "kilometers.")
        RentalFormToolTip.SetToolTip(AAAcheckbox, "To be checked if the customer is a member of AAA.")
        RentalFormToolTip.SetToolTip(Seniorcheckbox, "To be checked if the customer is a senior (65+).")

        RentalFormToolTip.SetToolTip(CalculateButton, "Verifies the customer's information and calculates " &
                                     "the total charges of the car rental.")
        RentalFormToolTip.SetToolTip(ClearButton, "Clears all textboxes and sets the buttons to their defaults.")
        RentalFormToolTip.SetToolTip(SummaryButton, "Displays the total amount of customers, car miles, " &
                                     "and charges recorded by this program.")
        RentalFormToolTip.SetToolTip(ExitButton, "Allows the user to exit the program.")
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        'Declaring variables
        Dim beginOdometer, endOdometer, day, dayCharge As Integer
        Dim mileage, mileageCost, discount, totalCost As Double


        'The inputs are checked using the ValidateUserEntry function, which returns true if the inputs
        'don't have errors and fills the odometer and day variables
        If ValidateUserEntry(beginOdometer, endOdometer, day) Then

            'Checks to see if the odometer values are in miles or kilometers, and converts the kilometers
            'to miles if checked.
            If MilesradioButton.Checked Then
                mileage = endOdometer - beginOdometer
            Else
                mileage = (endOdometer - beginOdometer) * 0.62  'converting from km to miles (1 mi = 0.62 km)
            End If


            'The mileage cost is based on the miles driven. From 0 - 200 mi there is no cost; from 201 - 500
            'each mile costs $0.12, and miles above 500 cost $0.10.
            If mileage > 500 Then
                mileageCost = (mileage - 500) * 0.1 'Applying the $0.10/mi cost
                'If there is more than 500 miles, the 200-500 mile charge will be automatically be applied
                mileageCost += 35.88

            ElseIf mileage > 200 Then
                mileageCost = (mileage - 200) * 0.12    'Applying the $0.12/mi cost

            Else    'If the mileage traveled was less than 200 mi, there are no mileage fees.
                mileageCost = 0
            End If

            dayCharge = day * 15    'There is a charge of $15 per day of car use

            'Depending on if the AAA or Senior checkboxes are checked, discounts will be recorded
            If AAAcheckbox.Checked Then
                discount = 0.05
            End If

            If Seniorcheckbox.Checked Then
                discount += 0.03
            End If

            'The discount percentage is multiplied by the total cost of the rental to find the discount amount
            discount = (dayCharge + mileageCost) * discount

            'The total cost is then the sum of the day charge and milage cost subtracted from the discount
            totalCost = (dayCharge + mileageCost - discount)

            'The cost values are formatted into currencies using the format function and "c" currency style
            'It will output a value with the format $##.##
            TotalMilesTextBox.Text = CStr(mileage) & " mi"
            MileageChargeTextBox.Text = Format(mileageCost, "c")
            DayChargeTextBox.Text = Format(dayCharge, "c")
            TotalDiscountTextBox.Text = Format(discount, "c")
            TotalChargeTextBox.Text = Format(totalCost, "c")

            'The totals are recorded in the CarCharges function so that they can be accessed in other subs
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

        'Declaring variables
        Dim totalsArray(2), displayString As String

        'The totalsArray retrieves the total customers, miles, and charges values from the CarCharges function
        totalsArray = CarCharges(0, 0, 0)

        'The totals values are labeled and formatted in the display string
        displayString = "Total Customers:" & vbTab & vbTab & totalsArray(0) & vbNewLine &
                        "Total Miles Driven:" & vbTab & vbTab & totalsArray(1) & " mi" & vbNewLine &
                        "Total Charges:" & vbTab & vbTab & "$" & totalsArray(2)

        'The summary totals are displayed to the user via a Message Box
        MsgBox(displayString,, "Detailed Summary")

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

        'If the user presses the exit button, the program verifies that the user wants to quit via a
        'Message box with yes and no buttons. If the user clicks yes, the program will close.
        Select Case MsgBox("Are you sure you want to exit?", vbYesNo, "Close")
            Case vbYes
                Me.Close()
            Case Else
                'If user doesn't want to exit, do nothing
        End Select

    End Sub

End Class
