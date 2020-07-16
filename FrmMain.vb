Imports System.IO
Imports System.Math
Imports System.Text
Public Class FrmMain
    Private Structure Times
        Public dayofyear As Date
        Public SunRise As Date
        Public Noon As Date
        Public Sunset As Date
    End Structure
    Dim MySun As New ClsSunrise
    Dim HaveValues As Boolean
    Dim fPath As String
    Dim TimeTable(365) As Times
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' populate city table
        PopulateCities(20)
        MySun.DaySavings = True
    End Sub

    Private Sub BtnCalculate_Click(sender As Object, e As EventArgs) Handles BtnCalculate.Click
        'MySun.TimeZone(1) = 1
        If MySun.TimeZone <> 0 And MySun.Latitude <> 0 And MySun.Longitude <> 0 Then
            HaveValues = True
        Else
            HaveValues = False
        End If
        If HaveValues Then
            MySun.CalculateSun()
            TxtSunrise.Text = FormatDateTime(MySun.Sunrise, DateFormat.LongTime)
            TxtSunset.Text = FormatDateTime(MySun.Sunset, DateFormat.LongTime)
            TxtHighNoon.Text = FormatDateTime(MySun.SolarNoon, DateFormat.LongTime)
        Else
            MsgBox("please select a City OR enter coordinates", vbOK)
        End If
    End Sub

    Private Sub PopulateCities(Count As Integer)
        ' this sub buils a list of cities for the City Selection dropdown box
        ' from the list supplied in CLSSunrise class
        Dim CityCount As Integer
        Dim CityName As String
        For CityCount = 0 To MySun.CityCount - 1
            CityName = MySun.CityName(CityCount)
            CmbCity.Items.Add(CityName)
        Next
    End Sub

    Private Sub CmbCity_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles CmbCity.SelectionChangeCommitted
        Dim CityName As String
        Dim CityIndex As Integer
        Dim CityInList As Boolean
        CityName = CmbCity.SelectedItem
        CityIndex = CmbCity.SelectedIndex
        CityInList = MySun.City(CityName)       'this call sets longitude and latidude in sunrise class
        If CityInList Then
            'MsgBox("city", 1, CityName)
            TxtLatitude.Text = MySun.Latitude
            TxtLongitude.Text = MySun.Longitude
        Else
            MsgBox("Please Enter Longitude and Latitude", vbYesNo)
        End If
    End Sub


    Private Sub TxtLongitude_TextChanged(sender As Object, e As EventArgs) Handles TxtLongitude.TextChanged
        'MsgBox("Changes", vbYes, "Longitude Box")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fSaveFile As New SaveFileDialog
        Dim Year As Date
        Dim YearStr As String
        'Dim fs As File
        Dim DayNumber As Integer
        'Dim OneLine As String
        fSaveFile.Filter = "(*.txt)|*.txt"
        fSaveFile.ShowDialog()
        fPath = fSaveFile.FileName
        FileOpen(1, fPath, OpenMode.Output)

        Year = DateTimePicker1.Value
        YearStr = "#1/1/" & Year.Year & "#"
        Year = YearStr
        Call BuildTimeTable(Year)
        For DayNumber = 0 To 365
            'OneLine = TimeTable(DayNumber).dayofyear
            WriteLine(1, Format(TimeTable(DayNumber).dayofyear, "MM/dd/yyyy"),
                      FormatDateTime(TimeTable(DayNumber).SunRise, DateFormat.ShortTime),
                      FormatDateTime(TimeTable(DayNumber).Sunset, DateFormat.ShortTime))
        Next
        FileClose(1)
        'Dim YearFile As FileIO
        'FileOpen(1, fPath, OpenAccess.Write)
        'WriteLine(1, 25)
        'FileClose(1)


    End Sub

    Private Sub BuildTimeTable(Year As Date)

        Dim Yearpart As Date
        Dim DayOfYear As Date
        Dim DayNumber As Integer


        Yearpart = Year
        DayOfYear = Yearpart
        For DayNumber = 0 To 365
            DayOfYear = DateAdd(DateInterval.Day, DayNumber, Yearpart)
            MySun.DateDay = DayOfYear
            MySun.CalculateSun()
            TimeTable(DayNumber).dayofyear = FormatDateTime(MySun.DateDay, DateFormat.ShortDate)
            TimeTable(DayNumber).SunRise = FormatDateTime(MySun.Sunrise, DateFormat.LongTime)
            TimeTable(DayNumber).Sunset = FormatDateTime(MySun.Sunset, DateFormat.LongTime)
            DateTimePicker1.Value = DayOfYear
            TxtSunrise.Text = FormatDateTime(MySun.Sunrise, DateFormat.LongTime)
            TxtSunset.Text = FormatDateTime(MySun.Sunset, DateFormat.LongTime)
            TxtHighNoon.Text = FormatDateTime(MySun.SolarNoon, DateFormat.LongTime)
        Next
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        MySun.DateDay = DateTimePicker1.Value.Date
    End Sub
End Class