'Project: Summer Olympic Recording System
'Assignment: Phase 3
'Programmer: Nguyen Quoc Trong Nghia - s3343711
'Created: April 30, 2013
'Purpose: Recording Athlete Information

Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

Public Class frmAddAthlete

    'Declare variables
    Dim sEventDate As String = Now.ToString("D")
    Const sEventVenue As String = "Flamengo Park"
    Const sEventName As String = "Long Jump"

    'Athlete Details
    Dim iId As Integer = 0
    Dim sFirstName As String = ""
    Dim sLastName As String = ""
    Dim sGender As String = ""
    Dim sCountry As String = ""
    Dim dPerformances(3) As Double
    Dim iPoints(3) As Integer
    Dim iAttempts() As Integer = {1, 2, 3, 4}
    Dim sStatusMessage As String
    Dim iPointScale As Integer
    Dim dStatistics(2) As Double

    'Others
    Dim bCheckNull As Boolean
    Dim iNumberOfAttempt As Integer
    Dim iLine As Integer
    Dim iSearchType As Integer = 0

    Private Sub frmAddAthlete_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Display Event Details
        txtEventDate.Text = sEventDate
        txtEventVenue.Text = sEventVenue
        txtEventName.Text = sEventName

        'Init default value
        cboCountry.SelectedIndex = 0
        cboSearchType.SelectedIndex = 0
        rdbMale.Select()

        For index = 0 To dStatistics.Count - 1
            dStatistics(index) = 0
        Next

        sStatusMessage = ""
        iPointScale = 0
        iNumberOfAttempt = 0

        bCheckNull = True
        doEnable()

    End Sub

    Private Sub MyTabs()

        'Declare variable for system initialization
        Me.tcMain = New TabControl()
        Me.tpAddAthlete = New TabPage()
        Me.tpAthleteDescription = New TabPage()
        Me.tpAthleteDetails = New TabPage()

        Me.tcMain.Controls.AddRange(New Control() {Me.tpAddAthlete, Me.tpAthleteDescription,Me.tpAthleteDetails})

        ' Selects tpAddathlete as main tab
        Me.tcMain.SelectedTab = tpAddAthlete

    End Sub

    Private Sub btnAddAthlete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAthlete.Click

        'call doValidation() to Checking all inputs from user
        doValidation()

    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click

        If Confirm("Clear") = True Then
            'Call doClear() to clear all textboxes
            doClear()
        End If

    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

        If Confirm("Exit") = True Then
            'Exit Program
            Me.Close()
        End If

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If txtDeleteAt.Text <> "" Then

            If Confirm("Delete") Then
                'Call doDelete() to delete specific item from lbAttempt
                doDelete()
            End If

        Else

            MsgBox("Please click an attempt number to delete")

        End If

    End Sub

    Private Sub btnDeleteAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteAll.Click

        If Confirm("Delete All") Then
            'Call doDeleteAll() to delete all detailed athlete performances and points 
            doDeleteAll()
        End If

    End Sub

    Private Sub btnPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Call calculatePoint() to calculate total point from athlete performance
        calculatePoint()
        doDisplayDetails()

    End Sub

    Private Sub btnClearPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearPoint.Click

        'Call clearPoint() to clear textboxes inside gbTotalPoint
        clearPoint()

    End Sub

    Private Sub btnStatistics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Call calculateStatistics() to calculate the highest, lowest and average performance of athlete
        calculateStatistics()
        doDisplayDetails()

    End Sub

    Private Sub btnClearStatistics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearStatistics.Click

        'Call clearStatistics() to clear all texboxes inside gbStatistics
        clearStatistics()

    End Sub

    Private Sub btnReAddPerformance_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReAddPerformance.Click

        'Call doReAdd() to add new performance
        doReAdd()

    End Sub

    Private Sub btnSaveFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFile.Click

        If Confirm("Save") = True Then
            'Call doSaveFile to save data to text file
            doSaveFile()
        End If

    End Sub

    Private Sub btnHtml_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHtml.Click

        'Call writeHTML to generate webpage
        writeHtml()

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        'Call doSearch to search athlete
        If validateInput() = True Then
            doSearch()
        End If

    End Sub

    Private Sub btnClearSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearSearch.Click

        If Confirm("Clear Search") = True Then
            'Call doClear() to clear all textboxes
            doClearSearch()
        End If

    End Sub

    Private Sub lbAttempt_MouseClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstAttempt.MouseClick

        'Display selected lbAttempt item
        txtDeleteAt.Text = lstAttempt.SelectedItem

    End Sub

    Private Sub cboSearchType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSearchType.SelectedIndexChanged

        'Clear all text for new Search Type
        txtMinValue.Clear()
        txtMaxValue.Clear()

        'Show textbox and label for input max value if search type is performance or point
        If cboSearchType.SelectedItem = "Performance" Or cboSearchType.SelectedItem = "Points" Then

            'Show textbox and label
            lblMax.Show()
            txtMaxValue.Show()
            lblMin.Text = "Min Value:"
            lblMax.Text = "Max Value:"

            'Assign value searchtype for later search purpose
            If cboSearchType.SelectedItem = "Performance" Then
                iSearchType = 3

            ElseIf cboSearchType.SelectedItem = "Points" Then
                iSearchType = 4

            End If

        Else 'Hide textbox and label for input max value if search type is performance or point

            'Hide textbox and label
            lblMax.Hide()
            txtMaxValue.Hide()
            lblMin.Text = cboSearchType.SelectedItem & ": "

            'Assign value searchtype for later search purpose
            If cboSearchType.SelectedItem = "ID" Then
                iSearchType = 0

            ElseIf cboSearchType.SelectedItem = "Last Name" Then
                iSearchType = 1

            End If

        End If

    End Sub

    Private Sub doClearAthleteDetails()

        'Clear all textboxes
        txtAthleteIdDetail.Clear()
        txtFirstNameDetails.Clear()
        txtLastNameDetails.Clear()
        txtGenderDetails.Clear()
        txtCountryDetails.Clear()
        txtAverageDetails.Clear()
        txtTotalDetails.Clear()
        txtStatusDetail.Clear()

    End Sub

    Private Sub doClearSearch()

        'Clear all listboxes result 
        lstResultId.Items.Clear()
        lstResultLastName.Items.Clear()
        lstResultFirstName.Items.Clear()
        lstResultPerformance.Items.Clear()
        lstResultPoint.Items.Clear()
        lstResultStatus.Items.Clear()

    End Sub

    Private Sub doClear()

        'Clear all Textboxes
        txtAthleteId.Clear()
        txtFirstName.Clear()
        txtLastName.Clear()
        txtPerformance1.Clear()
        txtPerformance3.Clear()
        txtPerformance2.Clear()
        txtPerformance4.Clear()

        'Select first item from dropdownlist
        cboCountry.SelectedIndex = 0
        'Select Male as default
        rdbMale.Select()

    End Sub

    Private Sub doSearch()

        'File Path
        Dim sFilename As String = Application.StartupPath & "\PerformanceDetails.txt"

        'File reading variable
        Dim srPerformanceFile As StreamReader

        ' Each line/record read from the file
        Dim sLine As String

        'Array to store data
        Dim sTemporaryData(99) As String

        Dim iCounter As Integer = 0

        If Not File.Exists(sFilename) Then
            'File not found
            MessageBox.Show("The file " & sFilename & " does not exist.")
        Else

            'Can proceed with opening the file because it exists
            srPerformanceFile = File.OpenText(sFilename)

            'read the file line by line
            sLine = srPerformanceFile.ReadLine

            'check if file contains data
            If sLine = "" Then
                MsgBox("File does not have any data")
            Else

                While sLine <> Nothing

                    'Store data to temporary string 
                    sTemporaryData(iCounter) = sLine

                    'Read next line in file
                    sLine = srPerformanceFile.ReadLine
                    iCounter += 1

                End While

                'Close stream reader
                srPerformanceFile.Close()

            End If

        End If

        'Clear all listboxes for new Search
        doClearSearch()

        If iSearchType = 0 Or iSearchType = 1 Then
            Dim iItemsCount As Integer = 0

            'Search
            For iIndex = 0 To iCounter - 1

                'Split data
                Dim sFields() As String = Split(sTemporaryData(iIndex), ",")

                'Check keyword with data of the file
                Dim checkData As Match = Regex.Match(txtMinValue.Text, sFields(iSearchType), RegexOptions.IgnoreCase)

                'Matched, add to listboxes
                If checkData.Success Then
                    lstResultId.Items.Add(sFields(0))
                    lstResultLastName.Items.Add(sFields(1))
                    lstResultFirstName.Items.Add(sFields(2))
                    lstResultPerformance.Items.Add(sFields(3))
                    lstResultPoint.Items.Add(sFields(4))
                    lstResultStatus.Items.Add(sFields(5))
                    iItemsCount += 1
                End If

            Next

            'No item added, display message
            If iItemsCount = 0 Then
                MsgBox("No data matched with search keyword")
            End If
        Else
            Dim iItemsCount As Integer = 0

            'Search
            For iIndex = 0 To iCounter - 1

                'Split data
                Dim sFields() As String = Split(sTemporaryData(iIndex), ",")

                'Take the value
                Dim dData As Double = Double.Parse(sFields(iSearchType))

                'Min, Max keywords
                Dim dMin As Double = Double.Parse(txtMinValue.Text)
                Dim dMax As Double = Double.Parse(txtMaxValue.Text)

                'In range, add to listboxes
                If dData >= dMin And dData <= dMax Then
                    lstResultId.Items.Add(sFields(0))
                    lstResultLastName.Items.Add(sFields(1))
                    lstResultFirstName.Items.Add(sFields(2))
                    lstResultPerformance.Items.Add(sFields(3))
                    lstResultPoint.Items.Add(sFields(4))
                    lstResultStatus.Items.Add(sFields(5))
                    iItemsCount += 1
                End If

            Next

            'No item added, display message
            If iItemsCount = 0 Then
                MsgBox("No data matched with search keyword")
            End If

        End If

    End Sub

    Private Sub writeHtml()

        'File Writing variable
        Dim swWriteFile As StreamWriter

        'File reading variable
        Dim srReadFile As StreamReader

        'Files' paths
        Dim sFileName As String = Application.StartupPath & "\CompetitionEvents.txt"
        Dim sDataWrite As String = Application.StartupPath & "CompetitionEvents.html"

        'Variable to store each line of the file
        Dim sLine As String

        'Array to store splitted data
        Dim sFields() As String

        Dim iNumberRecord As Integer

        Dim iMaxData As Integer = 3

        'Variables to store data
        Dim sEventDates(iMaxData) As String
        Dim sEventTimes(iMaxData) As String
        Dim sEventNames(iMaxData) As String
        Dim sEventLocations(iMaxData) As String

        iNumberRecord = 0

        If Not File.Exists(sFileName) Then
            'File not found
            MessageBox.Show("The file " & sFileName & " does not exist.")
        Else

            'Read text file
            srReadFile = File.OpenText(sFileName)

            'Store one line
            sLine = srReadFile.ReadLine

            While sLine <> Nothing
                'Split line to multiple data
                sFields = Split(sLine, ",")

                'Store splitted data to array
                sEventDates(iNumberRecord) = sFields(0)
                sEventTimes(iNumberRecord) = sFields(1)
                sEventNames(iNumberRecord) = sFields(2)
                sEventLocations(iNumberRecord) = sFields(3)

                iNumberRecord += 1

                'read new line
                sLine = srReadFile.ReadLine
            End While

            'Close file
            srReadFile.Close()

        End If

        'Write to html file
        swWriteFile = New StreamWriter(sDataWrite)

        'Write html content
        swWriteFile.WriteLine("<!DOCTYPE html>")

        swWriteFile.WriteLine("<html>")
        swWriteFile.WriteLine("<HEAD> ")
        swWriteFile.WriteLine("<TITLE>Competition Events</TITLE>")
        swWriteFile.WriteLine(" </HEAD>")
        swWriteFile.WriteLine("<BODY>")
        swWriteFile.WriteLine("<div align=""center"" >")
        swWriteFile.WriteLine("<h1>Competition Event</h1>")
        swWriteFile.WriteLine("<h2>Event Records</h2>")

        'Table content
        swWriteFile.WriteLine("<table border=""1"">")
        swWriteFile.WriteLine("<tr>")
        swWriteFile.WriteLine("<th>Event Dates</th>")
        swWriteFile.WriteLine("<th>Event Times</th>")
        swWriteFile.WriteLine("<th>Event Names</th>")
        swWriteFile.WriteLine("<th>Event Locations</th>")
        swWriteFile.WriteLine("</tr>")

        'Table data
        For iIndex = 0 To sEventDates.Count - 1
            swWriteFile.WriteLine("<tr>")
            swWriteFile.WriteLine("<td>" & sEventDates(iIndex) & "</td>")
            swWriteFile.WriteLine("<td>" & sEventTimes(iIndex) & "</td>")
            swWriteFile.WriteLine("<td>" & sEventNames(iIndex) & "</td>")
            swWriteFile.WriteLine("<td>" & sEventLocations(iIndex) & "</td>")
            swWriteFile.WriteLine("</tr>")
        Next

        swWriteFile.WriteLine("</table>")
        swWriteFile.WriteLine("</div>")
        swWriteFile.WriteLine("</BODY>")
        swWriteFile.WriteLine("</HTML>")

        'Close file
        swWriteFile.Close()

        'Open Browser
        Process.Start(sDataWrite)

    End Sub

    Private Sub doSaveFile()

        If bCheckNull <> True Then
            'File Path
            Dim sFilename As String = Application.StartupPath & "\PerformanceDetails.txt"

            'File reading variable
            Dim srPerformanceFile As StreamReader

            'File writing variable
            Dim swPerformanceFile As StreamWriter

            ' Each line/record read from the file
            Dim sLine As String

            'Declare data to write
            Dim sDataWrite As String = iId & "," & sLastName & "," & sFirstName & "," &
                                        dStatistics(2) & "," & iPointScale &
                                            "," & sStatusMessage & vbCrLf

            'Temporary string to store old data
            Dim sTemporaryData As String = ""

            If Not File.Exists(sFilename) Then
                'File not found
                MessageBox.Show("The file " & sFilename & " does not exist.")
            Else

                'Can proceed with opening the file because it exists
                srPerformanceFile = File.OpenText(sFilename)

                'read the file line by line
                sLine = srPerformanceFile.ReadLine
                iLine = 0

                While sLine <> Nothing
                    'Store data to temporary string 
                    sTemporaryData += sLine & vbCrLf

                    'Read next line in file
                    sLine = srPerformanceFile.ReadLine

                    iLine += 1
                End While

                'Close stream reader
                srPerformanceFile.Close()

                'Open file to write
                swPerformanceFile = New StreamWriter(sFilename)

                'Write to file
                sTemporaryData += sDataWrite
                swPerformanceFile.WriteLine(sTemporaryData)

                'Close stream writer
                swPerformanceFile.Close()

                'Confirm Message
                MsgBox("Save Completed")

            End If

        End If

    End Sub

    Private Sub doDisplayDetails()

        If bCheckNull <> True Then

            'Display Athlete Details Tab
            txtAthleteIdDetail.Text = iId
            txtLastNameDetails.Text = sLastName
            txtFirstNameDetails.Text = sFirstName
            txtGenderDetails.Text = sGender
            txtCountryDetails.Text = sCountry
            txtAverageDetails.Text = Format(calculateStatistics, "0.00")
            txtStatusDetail.Text = calculatePoint()
            txtTotalDetails.Text = iPointScale

        End If

    End Sub

    Private Sub doReAdd()

        'Declare vairable for checking user input
        Dim bCheckValid As Boolean = validateDecimal(txtReAddPerformance.Text)

        'Invalid input, display error message
        If bCheckValid <> True Then
            MsgBox("Invalid Performance", MsgBoxStyle.OkOnly, "Error")
        Else

            'Valid input, add item to listboxes
            'Check missing performance and re-add
            For iIndex As Integer = 0 To iAttempts.Count - 1

                If dPerformances(iIndex) = 0 And iPoints(iIndex) = 0 Then
                    dPerformances(iIndex) = Double.Parse(txtReAddPerformance.Text)
                    iPoints(iIndex) = checkPoint(dPerformances(iIndex))
                    Exit For
                End If

            Next

            'Clear listboxes before adding items
            lstAttempt.Items.Clear()
            lstPerformance.Items.Clear()
            lstPoint.Items.Clear()

            'Add items to listboxes based on 3 array
            For iIndex As Integer = 0 To iAttempts.Count - 1

                If dPerformances(iIndex) <> 0 And iPoints(iIndex) <> 0 Then
                    lstAttempt.Items.Add(iAttempts(iIndex))
                    lstPerformance.Items.Add(dPerformances(iIndex))
                    lstPoint.Items.Add(iPoints(iIndex))
                End If

            Next

            iNumberOfAttempt += 1

            'Enable button
            doEnable()

            'Update 3 listboxes
            lstAttempt.Update()
            lstPerformance.Update()
            lstPoint.Update()

            'Clear user input
            txtReAddPerformance.Clear()

            'Update Points and Statistics Summary
            calculatePoint()
            calculateStatistics()
            doDisplayDetails()

        End If

    End Sub

    Private Sub doDelete()

        'Check iNumberOfAttempt, =0, disable buttons
        If iNumberOfAttempt <> 0 Then
            'Store the deleted item's position
            Dim iPosition As Integer = Integer.Parse(lstAttempt.SelectedItem) - 1
            dPerformances(iPosition) = 0
            iPoints(iPosition) = 0

            'Clear listboxes before adding items
            lstAttempt.Items.Clear()
            lstPerformance.Items.Clear()
            lstPoint.Items.Clear()

            'Add items to listboxes
            For iIndex As Integer = 0 To iAttempts.Count - 1

                If dPerformances(iIndex) <> 0 And iPoints(iIndex) <> 0 Then
                    lstAttempt.Items.Add(iAttempts(iIndex))
                    lstPerformance.Items.Add(dPerformances(iIndex))
                    lstPoint.Items.Add(iPoints(iIndex))
                End If

            Next

            iNumberOfAttempt -= 1

            'iNumberOfAttempt =0, disable buttons, clear Points and Statistics Summary
            If iNumberOfAttempt = 0 Then
                bCheckNull = True
                doEnable()
                clearPoint()
                clearStatistics()

                'Clear Athlete Details
                doClearAthleteDetails()

            Else

                doEnable()

                'Update listboxes
                lstAttempt.Update()
                lstPerformance.Update()
                lstPoint.Update()

                'Clear user input
                txtDeleteAt.Clear()

                'Update Points and Statistics Summary
                calculatePoint()
                calculateStatistics()
                doDisplayDetails()

            End If

        Else

            bCheckNull = True
            doEnable()

            'Clear Athlete Details
            doClearAthleteDetails()

        End If

    End Sub

    Private Sub doDeleteAll()

        'Clear all items from listboxes
        lstAttempt.Items.Clear()
        lstPerformance.Items.Clear()
        lstPoint.Items.Clear()

        'Clear all textboxes
        txtDeleteAt.Clear()
        txtAthleteId2.Clear()

        'Clear all Points and Statistics Summary
        clearPoint()
        clearStatistics()

        'Disable buttons
        bCheckNull = True
        iNumberOfAttempt = 0
        doEnable()

        'Clear Athlete Details
        doClearAthleteDetails()

    End Sub

    Private Sub doAdd()

        If bCheckNull = False Then
            'Display Athlete ID
            txtAthleteId2.Text = iId

            'Adding Attempt number, Performance, Point to lbAttempt, lbPerformance, lbPoint accordingly
            For iIndex As Integer = 0 To iAttempts.Count - 1
                lstAttempt.Items.Add(iAttempts(iIndex))
                lstPerformance.Items.Add(dPerformances(iIndex))
                lstPoint.Items.Add(iPoints(iIndex))
                iNumberOfAttempt += 1
            Next

            'Update all listboxes after adding items
            lstAttempt.Update()
            lstPerformance.Update()
            lstPoint.Update()

            'Calculate Points and Statistics Summary
            calculatePoint()
            calculateStatistics()
            doDisplayDetails()

            'Enable all buttons of tbAthleteDescription
            doEnable()

        End If

    End Sub

    Private Sub doValidation()

        'Declare variables       
        Dim sError As String = ""
        Dim sMessage As String = ""

        'Declare regular expression
        Dim sRegexIdPattern As String = "^([0-9]{5,5})$"
        Dim regexStringPattern As String = "([A-Za-z ]+)$"

        'All fields must be filled with input
        If txtAthleteId.Text = "" Or
            txtFirstName.Text = "" Or txtLastName.Text = "" Or cboCountry.SelectedItem = "" Or
            txtPerformance1.Text = "" Or txtPerformance3.Text = "" Or
            txtPerformance2.Text = "" Or txtPerformance4.Text = "" Then
            sError += "All fields are required." & vbCrLf
        Else

            'Athlete Id Validation (Error will display if not 5 integer numbers format)
            Dim checkId As Match = Regex.Match(txtAthleteId.Text, _
                        sRegexIdPattern, _
                    RegexOptions.IgnoreCase)

            'Validation not meet, add error message to error list
            If checkId.Success <> True Then
                sError += "Invalid ID. Must be 5 digit" & vbCrLf
                txtAthleteId.Clear()
            Else

                iId = Integer.Parse(txtAthleteId.Text)

                'Athlete Id Must be between 10000 and 99999
                'Validation not meet, add error message to error list
                If iId < 10000 Or iId > 99999 Then
                    sError += "Invalid ID. Must be between 10000 and 99999"
                End If

            End If

            'First Name validation (Error will display if not alphabetical format)
            Dim checkFirstName As Match = Regex.Match(txtFirstName.Text, _
                        regexStringPattern, _
                    RegexOptions.IgnoreCase)

            'Validation not meet, add error message to error list
            If checkFirstName.Success <> True Then
                sError += "First Name contains characters only" & vbCrLf
                txtFirstName.Clear()
            End If

            'Last Name validation (Error will display if not alphabetical format)
            Dim checkLastName As Match = Regex.Match(txtLastName.Text, _
                        regexStringPattern, _
                    RegexOptions.IgnoreCase)

            'Validation not meet, add error message to error list
            If checkLastName.Success <> True Then
                sError += "Last Name contains characters only" & vbCrLf
                txtLastName.Clear()
            End If

            'Attemp Performance Validation Checking variables
            Dim bCheckPerformance1 As Boolean = validateDecimal(txtPerformance1.Text)
            Dim bCheckPerformance2 As Boolean = validateDecimal(txtPerformance2.Text)
            Dim bCheckPerformance3 As Boolean = validateDecimal(txtPerformance3.Text)
            Dim bCheckPerformance4 As Boolean = validateDecimal(txtPerformance4.Text)

            'If all performances are in correct format, parse them to dPerformances()
            If bCheckPerformance1 = True And bCheckPerformance2 = True And
                bCheckPerformance3 = True And bCheckPerformance4 = True Then

                'Passed validation, parse all textbox into array of Double
                dPerformances(0) = Double.Parse(txtPerformance1.Text)
                dPerformances(1) = Double.Parse(txtPerformance2.Text)
                dPerformances(2) = Double.Parse(txtPerformance3.Text)
                dPerformances(3) = Double.Parse(txtPerformance4.Text)

                'Performance cannot be 0
                If dPerformances(0) <> 0 Or dPerformances(1) <> 0 Or
                    dPerformances(2) <> 0 Or dPerformances(3) <> 0 Then

                    'Calculate Scale Point based on Performance
                    For iIndex As Integer = 0 To dPerformances.Length - 1
                        iPoints(iIndex) = checkPoint(dPerformances(iIndex))
                    Next

                Else

                    For iIndex = 0 To 3
                        'If Performance  is 0, add error message to error list
                        If dPerformances(iIndex) = 0 Then
                            sError += "Invalid Performance" & iIndex & ". Performance cannot be 0" & vbCrLf
                        End If
                    Next
                    
                End If

            Else

                'Check performance 1, not pass all validation, add error message to error list
                If bCheckPerformance1 <> True Then
                    sError += "Invalid Performance 1. Must be positive and cannot exceed more than 2 decimal places" & vbCrLf
                    txtPerformance1.Clear()
                End If

                'Check performance 2, not pass all validation, add error message to error list
                If bCheckPerformance2 <> True Then
                    sError += "Invalid Performance 2. Must be positive and cannot exceed more than 2 decimal places" & vbCrLf
                    txtPerformance2.Clear()
                End If

                'Check performance 3, not pass all validation, add error message to error list
                If bCheckPerformance3 <> True Then
                    sError += "Invalid Performance 3. Must be positive and cannot exceed more than 2 decimal places" & vbCrLf
                    txtPerformance3.Clear()
                End If

                'Check performance 4, not pass all validation, add error message to error list
                If bCheckPerformance4 <> True Then
                    sError += "Invalid Performance 4. Must be positive and cannot exceed more than 2 decimal places" & vbCrLf
                    txtPerformance4.Clear()
                End If

            End If

        End If

        'No errors so far
        If sError = "" Then

            'Display Message confirmation that Data is saved
            sFirstName = txtFirstName.Text
            sLastName = txtLastName.Text
            sCountry = cboCountry.SelectedItem

            If rdbFemale.Checked Then
                sGender = "Female"
            Else
                sGender = "Male"
            End If

            MsgBox("Data Saved", MsgBoxStyle.Information, "Save Completed")

            'Add all user inputs to tbAthleteDescription
            bCheckNull = False
            doAdd()

            'Display Athlete Details tab
            doDisplayDetails()
            Me.tcMain.SelectedTab = tpAthleteDetails

        Else
            'Display Errors from user
            MsgBox(sError, MsgBoxStyle.OkOnly, "Error List")
        End If

    End Sub

    Private Sub doEnable()

        'if bCheckNull equals to true, disbable all buttons
        If bCheckNull = True Then

            btnClearPoint.Enabled = False
            btnDelete.Enabled = False
            btnDeleteAll.Enabled = False
            btnClearStatistics.Enabled = False
            btnAddAthlete.Enabled = True
            btnReAddPerformance.Enabled = False
            btnPoint.Enabled = False
            btnStatistics.Enabled = False
            btnSaveFile.Enabled = False

        Else

            'if bCheckNull equals to false, enable all buttons
            btnClearPoint.Enabled = True
            btnDelete.Enabled = True
            btnDeleteAll.Enabled = True
            btnClearStatistics.Enabled = True
            btnAddAthlete.Enabled = False
            btnPoint.Enabled = True
            btnStatistics.Enabled = True
            btnSaveFile.Enabled = True

            'iNumberOfAttempt =4, disable this button, else enable
            If iNumberOfAttempt = 4 Then
                btnReAddPerformance.Enabled = False
            Else
                btnReAddPerformance.Enabled = True
            End If

        End If

    End Sub

    Private Sub clearStatistics()

        'Clear all text from gbStatistics textboxes
        txtHighest.Clear()
        txtLowest.Clear()
        txtAverage.Clear()

    End Sub

    Private Sub clearPoint()

        'Clear all text from gbTotalPoint texboxes
        txtTotalPoint.Clear()
        txtStatus.Clear()

    End Sub

    Private Function calculatePoint() As String

        If bCheckNull = False Then
            'Clear texboxes before displaying summary
            txtTotalPoint.Clear()
            txtStatus.Clear()

            'Calculate total points from lbPoint
            iPointScale = 0

            For iIndex As Integer = 0 To lstPoint.Items.Count - 1
                iPointScale += lstPoint.Items(iIndex)
            Next

            'If total point >=15, qualified, else unqualified
            If iPointScale >= 15 Then
                sStatusMessage = "Qualified for the next round"
            Else
                sStatusMessage = "Unqualified for the next round"
            End If

            'Display result to the screen
            txtTotalPoint.Text = iPointScale
            txtStatus.Text = sStatusMessage

        End If

        Return sStatusMessage

    End Function

    Private Function calculateStatistics() As Double

        If bCheckNull = False Then
            'Declare variables for calculate Statistics Summary
            Dim dHighest As Double = lstPerformance.Items(0)
            Dim dLowest As Double = lstPerformance.Items(0)
            Dim dAverage As Double = 0

            'Clear all texboxes before displaying summary
            txtHighest.Clear()
            txtLowest.Clear()
            txtAverage.Clear()

            For iIndex As Integer = 0 To lstPerformance.Items.Count - 1

                'Calculate Highest
                If dHighest < lstPerformance.Items(iIndex) Then
                    dHighest = lstPerformance.Items(iIndex)
                End If

                'Calculate Lowest
                If dLowest > lstPerformance.Items(iIndex) Then
                    dLowest = lstPerformance.Items(iIndex)
                End If

                'Calculate Average
                dAverage += lstPerformance.Items(iIndex)

            Next

            'Calculate Average            
            dAverage = dAverage / (lstPerformance.Items.Count)

            'Add data to array
            dStatistics(0) = Format(dHighest, "0.00")
            dStatistics(1) = Format(dLowest, "0.00")
            dStatistics(2) = Format(dAverage, "0.00")

            'Display to the Screen
            txtHighest.Text = dStatistics(0)
            txtLowest.Text = dStatistics(1)
            txtAverage.Text = dStatistics(2)

        End If

        Return dStatistics(2)

    End Function

    Private Function validateDecimal(ByVal sValidate As String) As Boolean

        'Declare variables for validation
        Dim sRegexDecimalPattern As String = "^[+0-9]+(\.[0-9]{1,2})?$"
        Dim checkDecimal As Match = Regex.Match(sValidate, sRegexDecimalPattern, RegexOptions.IgnoreCase)

        'If verify failed, return false
        If checkDecimal.Success <> True Then
            Return False
        Else

            'If verify passed, parse sValidate to double
            Dim dParseDouble As Double = Double.Parse(sValidate)

            'If parsed value equals 0, return false
            If dParseDouble = 0 Then
                Return False
            Else
                'If parsed value not equals 0, return true
                Return True
            End If

        End If

    End Function

    Private Function validatePoint(ByVal sValidate As String) As Boolean

        'Declare variables for validation
        Dim sRegexPointPattern As String = "^[0-9]+$"
        Dim checkPoint As Match = Regex.Match(sValidate, sRegexPointPattern, RegexOptions.IgnoreCase)

        'If verify failed, return false
        If checkPoint.Success <> True Then
            Return False
        Else

            'If verify passed, parse sValidate to integer
            Dim iParseInteger As Integer = Integer.Parse(sValidate)

            'If parsed value equals 0, return false
            If iParseInteger = 0 Then
                Return False
            Else
                'If parsed value not equals 0, return true
                Return True
            End If

        End If

    End Function

    Private Function checkPoint(ByVal dInput As Double) As Integer

        If dInput <= 0.5 Then

            'Return 1 if dInput <=0.5
            Return 1

        ElseIf dInput <= 1.0 Then

            'Return 2 if dInput <=1.0
            Return 2

        ElseIf dInput <= 1.5 Then

            'Return 4 if dInput <=1.5
            Return 4

        Else

            'Return 6 if dInput >1.5
            Return 6

        End If

    End Function

    Private Function Confirm(ByVal sType As String) As Boolean

        'Declare a message confirmation
        Dim response As MsgBoxResult

        If sType = "Exit" Then
            'Exit confirmation
            response = MsgBox("Do you want to Exit", MsgBoxStyle.YesNo, "Exit Program")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If

        ElseIf sType = "Clear" Then

            'Clear Confirmation
            response = MsgBox("Do you want to Clear", MsgBoxStyle.YesNo, "Clear All Details")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If

        ElseIf sType = "Clear Search" Then

            'Clear Confirmation
            response = MsgBox("Do you want to Clear Search Result", MsgBoxStyle.YesNo, "Clear Search Result")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If

        ElseIf sType = "Delete All" Then

            'Delete All Confirmation
            response = MsgBox("Do you want to Delete All", MsgBoxStyle.YesNo, "Delete All Details")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If

        ElseIf sType = "Save" Then

            'Delete All Confirmation
            response = MsgBox("Do you want to Save to Text file", MsgBoxStyle.YesNo, "Save Data")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If
        Else

            'Delete Confirmation
            response = MsgBox("Do you want to Delete", MsgBoxStyle.YesNo, "Delete One Item")

            If response = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If

        End If

    End Function

    Private Function validateInput() As Boolean

        'boolean to check input
        Dim bCheck As Boolean = True

        'regular expression for checking input
        Dim sRegexIdPattern As String = "^([0-9]{5,5})$"
        Dim sRegexPointPattern As String = "^([0-9]+)$"
        Dim regexStringPattern As String = "([A-Za-z ]+)$"

        'Search ID
        If iSearchType = 0 Then

            'Check blank
            If txtMinValue.Text = "" Then
                MsgBox("ID cannot be blanked")
                bCheck = False
            Else

                'Athlete Id Validation (Error will display if not 5 integer numbers format)
                Dim checkId As Match = Regex.Match(txtMinValue.Text, _
                            sRegexIdPattern, _
                        RegexOptions.IgnoreCase)

                'Validation not meet, add error message to error list
                If checkId.Success <> True Then
                    MsgBox("Invalid ID")
                    txtMinValue.Clear()
                    bCheck = False
                Else

                    iId = Integer.Parse(txtMinValue.Text)

                    'Athlete Id Must be between 10000 and 99999
                    'Validation not meet, add error message to error list
                    If iId < 10000 Or iId > 99999 Then
                        MsgBox("ID Must be between 10000 and 99999")
                        txtMinValue.Clear()
                        bCheck = False
                    End If

                End If

            End If

            'Search Last Name
        ElseIf iSearchType = 1 Then

            'Check blank
            If txtMinValue.Text = "" Then
                MsgBox("Last Name cannot be blanked")
                bCheck = False
            Else

                'Last Name validation (Error will display if not alphabetical format)
                Dim checkLastName As Match = Regex.Match(txtMinValue.Text, _
                            regexStringPattern, _
                        RegexOptions.IgnoreCase)

                'Validation not meet, add error message to error list
                If checkLastName.Success <> True Then
                    MsgBox("Invalid Last Name")
                    txtMinValue.Clear()
                    bCheck = False
                End If

            End If

            'Search Performance
        ElseIf iSearchType = 3 Then

            'Check blank
            If txtMinValue.Text = "" Or txtMaxValue.Text = "" Then
                MsgBox("All fields required")
                bCheck = False
            Else

                'Check Min, Max inputs
                If validateDecimal(txtMinValue.Text) = True And validateDecimal(txtMaxValue.Text) = True Then

                    'Inputs passed all validation
                    If Double.Parse(txtMinValue.Text) < Double.Parse(txtMaxValue.Text) Then
                        bCheck = True
                    Else
                        'Min> Max, display error message
                        MsgBox("Min Value cannot greater than Max value")
                        bCheck = False
                    End If

                Else

                    'Validate min input
                    If validateDecimal(txtMinValue.Text) <> True Then
                        'Invalid Min input
                        MsgBox("Invalid Min Value")
                    End If

                    'Validate max input
                    If validateDecimal(txtMaxValue.Text) <> True Then
                        'Invalid Min input
                        MsgBox("Invalid Max Value")
                    End If

                    bCheck = False

                End If

            End If

            'Search Point
        ElseIf iSearchType = 4 Then

            'Check blank
            If txtMinValue.Text = "" Or txtMaxValue.Text = "" Then
                MsgBox("All fields required")
                bCheck = False
            Else

                If validatePoint(txtMinValue.Text) = True And validatePoint(txtMaxValue.Text) = True Then

                    'Inputs passed all validations
                    If Integer.Parse(txtMinValue.Text) < Integer.Parse(txtMaxValue.Text) Then
                        bCheck = True
                    Else
                        'Min > max, display error message
                        MsgBox("Min Value cannot greater than Max value")
                        bCheck = False
                    End If

                Else

                    'Validate min, max inputs
                    If validatePoint(txtMinValue.Text) <> True And validatePoint(txtMaxValue.Text) <> True Then
                        'min, max did not pass
                        MsgBox("Invalid Min, Max Value")
                    Else

                        'Validate max input
                        If validatePoint(txtMaxValue.Text) <> True Then
                            'Invalid max value
                            MsgBox("Invalid Max Value")
                        End If

                        'Validate min input
                        If validatePoint(txtMinValue.Text) <> True Then
                            'Invalid min value
                            MsgBox("Invalid Min Value")
                        End If

                    End If

                    bCheck = False

                End If

            End If

        End If

        Return bCheck

    End Function

End Class
