Imports Microsoft.Office.Core
Imports Microsoft.Win32
Imports Newtonsoft
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Windows
Imports System.Windows.Forms
Imports System.Drawing
Imports stdole
Imports System.Collections

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility
    Implements Microsoft.Office.Core.IRibbonControl

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("TrelloCalendar.Ribbon1.xml")
    End Function

    'GLOBAL CONSTANTS
    Const CalendarName As String = "Trello Calendar"
    Const TrelloBase As String = "https://trello.com/"
    Const TrelloKey As String = "c220897e88bf3b1bcc450bef48c63254"
    Const TrelloName As String = "Trello+Outlook+Addon"
    Const TrelloVersion As String = "1"
    Const TrelloRegKey As String = "SOFTWARE\Trello Outlook Addon"
    Const TrelloMappingRegKey As String = "SOFTWARE\Trello Outlook Addon\Mapping"
    Const cbHeight As Integer = 20
    Const cbWidth As Integer = 210
    Const cbMarginTop As Integer = 4
    Const cbMarginLeft As Integer = 40
    Const cbStart As Integer = 60

    'PRIVATE GLOBAL VARIABLES
    Private ribbon As Office.IRibbonUI
    Private configOpen As Boolean = False
    Private authOpen As Boolean = False
    Private container As New Form
    Private wb As New WebBrowser
    Private token As String = ""
    Private objCalendar As Outlook.MAPIFolder
    Private config As New JObject
    Private explorer As Outlook.Explorer

#Region "Ribbon Callbacks"

    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub btnDeauthorise_Click(ByRef control)
        DeAuthoriseApplication()
    End Sub

    Public Sub btnTrello_Click(ByRef control)
        System.Diagnostics.Process.Start("https://www.trello.com")
    End Sub

    Public Sub btnSync_Click(ByRef control)
        SyncFromTrello()
    End Sub

    Public Sub btnAuth_Click(ByRef control)
        ShowAuthorisePrompt()
    End Sub

    Public Sub btnConfig_Click(ByRef control)
        CreateConfig()
    End Sub

    Public Sub btnFix_Click(ByRef control)
        FixMe()
    End Sub

#End Region



#Region "Public Functions"
    Public Sub StartUp()
        CreateRequiredRegistryKeys()
        RefreshToken()
        explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        CreateTrelloCalendar()
        If CheckEnvironment(False, False, True, True) Then
            AttachContainerHandlers()
            AttachAppointmentHandlers()
        Else
            ShowError("Trello Outlook Addon was not loaded", "Internal Error")
        End If
    End Sub
#End Region



#Region "Trello Functions"

    Private Function BuildUrl(ByVal action As String, ByVal params As String)
        'BUILD URL FOR TRELLO API CALL
        Dim url As String = ""
        If Not String.IsNullOrEmpty(action) Then
            url = TrelloBase & TrelloVersion & "/" & action
            If Not String.IsNullOrEmpty(params) Then
                url = url & "?" & params & "&"
            Else
                url = url & "?"
            End If
            url = url & "key=" & TrelloKey & "&token=" & token
        End If
        Return url
    End Function

    Private Overloads Sub CallTrello(ByVal action As String, ByVal params As String, ByRef content As JObject)
        Dim client As New WebClient()
        Dim url As String = BuildUrl(action, params)
        Dim ret As String = ""
        If Not String.IsNullOrEmpty(url) Then
            ret = client.DownloadString(url)
            content = JsonConvert.DeserializeObject(ret)
        End If
    End Sub

    Private Overloads Sub CallTrello(ByVal action As String, ByVal params As String, ByRef content As JArray)
        Dim client As New WebClient()
        Dim url As String = BuildUrl(action, params)
        Dim ret As String = ""
        If Not String.IsNullOrEmpty(url) Then
            ret = client.DownloadString(url)
            content = JsonConvert.DeserializeObject(ret)
        End If
    End Sub

#End Region



#Region "Registry Functions"

    Private Sub SplitSubKey(ByVal subkey As String, ByRef root As String, ByRef key As String)
        Dim temp As Array
        Dim ret As String = ""
        temp = subkey.Split("\")
        For i As Integer = 0 To temp.Length - 1
            If i = temp.Length - 1 Then
                root = ret
                key = temp(i)
            Else
                If ret.Length > 0 Then
                    ret = ret & "\" & temp(i)
                Else
                    ret = temp(i)
                End If
            End If
        Next
    End Sub

    Public Sub CreateSubKey(ByVal SubKey As String, ByVal KeyName As String)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        regKey.CreateSubKey(KeyName)
        regKey.Close()
    End Sub

    Public Sub DeleteSubKey(ByVal SubKey As String, ByVal KeyName As String)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        regKey.DeleteSubKey(KeyName, False)
        regKey.Close()
    End Sub

    Public Function ReadRegKeyStr(ByVal SubKey As String, ByVal KeyName As String) As String
        Dim regKey As RegistryKey
        Dim value As String
        regKey = OpenSubKey(SubKey)
        value = regKey.GetValue(KeyName, "")
        regKey.Close()
        Return value
    End Function

    Public Overloads Function ReadRegKeyInt(ByVal SubKey As String, ByVal KeyName As String) As Integer
        Dim regKey As RegistryKey
        Dim value As Integer
        regKey = OpenSubKey(SubKey)
        value = regKey.GetValue(KeyName, 0)
        regKey.Close()
        Return value
    End Function

    Public Overloads Function ReadRegKeyDec(ByVal SubKey As String, ByVal KeyName As String) As Decimal
        Dim regKey As RegistryKey
        Dim value As Decimal
        regKey = OpenSubKey(SubKey)
        value = regKey.GetValue(KeyName, 0.0)
        regKey.Close()
        Return value
    End Function

    Public Overloads Sub WriteRegKey(ByVal SubKey As String, ByVal KeyName As String, ByVal Value As String)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        regKey.SetValue(KeyName, Value)
        regKey.Close()
    End Sub

    Public Overloads Sub WriteRegKey(ByVal SubKey As String, ByVal KeyName As String, ByVal Value As Integer)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        regKey.SetValue(KeyName, Value)
        regKey.Close()
    End Sub

    Public Overloads Sub WriteRegKey(ByVal SubKey As String, ByVal KeyName As String, ByVal Value As Decimal)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        regKey.SetValue(KeyName, Value)
        regKey.Close()
    End Sub

    Public Sub DeleteRegKey(ByVal SubKey As String, ByVal KeyName As String)
        Dim regKey As RegistryKey
        regKey = OpenSubKey(SubKey)
        If Not regKey.GetValue(KeyName, Nothing) Is Nothing Then
            regKey.DeleteValue(KeyName)
        End If
        regKey.Close()
    End Sub

    Private Function OpenSubKey(ByVal SubKey As String) As RegistryKey
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey(SubKey, True)
        If (regKey Is Nothing) Then
            Dim Root As String = ""
            Dim Key As String = ""
            SplitSubKey(SubKey, Root, Key)
            CreateSubKey(Root, Key)
            regKey = Registry.CurrentUser.OpenSubKey(SubKey, True)
        End If
        If (regKey Is Nothing) Then
            Throw New ApplicationException("Error writing registry key in: '" & SubKey & "'")
        End If
        Return regKey
    End Function

    Private Function FindRegKeyByTrelloId(ByVal TrelloId As String) As JObject
        Dim Value As New JObject
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey(TrelloMappingRegKey, True)
        Value.RemoveAll()
        For Each regValue In regKey.GetValueNames()
            If Not String.IsNullOrEmpty(regValue) Then
                Dim JSON As JObject = JsonConvert.DeserializeObject(regKey.GetValue(regValue))
                If JSON.HasValues Then
                    If ((JSON("id") IsNot Nothing) AndAlso (JSON("id").ToString = TrelloId)) Then
                        Value = JSON
                    End If
                End If
            End If
        Next
        Return Value
    End Function

#End Region



#Region "Private Functions"

    Private Sub WriteMapping(ByVal OutlookApptGlobalId As String, JSONCard As String)
        WriteRegKey(TrelloMappingRegKey, OutlookApptGlobalId, JSONCard)
    End Sub

    Private Sub SaveConfig(sender As Object, e As System.EventArgs)
        If CheckEnvironment(True, False) Then
            Dim Push As Boolean = False
            Dim Assigned As Boolean = False
            Dim config As New JObject
            Dim boards As New JArray
            For Each ctrl As Control In container.Controls
                If TypeOf ctrl Is CheckBox Then
                    If ctrl.Name.IndexOf("cb-") > -1 AndAlso DirectCast(ctrl, CheckBox).Checked Then
                        Select Case ctrl.Name.ToLower
                            Case "cb-pushtrello"
                                Push = True
                            Case "cb-assignedtome"
                                Assigned = True
                            Case Else
                                Dim board As New JObject
                                board.Add(New JProperty("id", ctrl.Name.Replace("cb-", "")))
                                board.Add(New JProperty("name", ctrl.Text))
                                boards.Add(board)
                        End Select
                    End If
                End If
            Next
            config.Add(New JProperty("boards", boards))
            config.Add(New JProperty("push", Push))
            config.Add(New JProperty("assigned", Assigned))
            WriteRegKey(TrelloRegKey, "config", config.ToString())
            GetConfig()
            CreateTrelloCalendar()
            ShowInformation("Configuration Saved!", "Configuration")
            CancelClose(sender, Nothing)
        End If
    End Sub

    Private Sub RefreshToken()
        token = ReadRegKeyStr(TrelloRegKey, "token")
    End Sub

    Public Function GetImage(ByVal control As IRibbonControl) As IPictureDisp
        Dim id As String = control.Id.ToUpper
        Select Case id
            Case "BTNTRELLO"
                Return ImageConverter.Convert(My.Resources.trello.ToBitmap())
            Case "BTNFIX"
                Return ImageConverter.Convert(My.Resources.rescue.ToBitmap())
        End Select
        Return Nothing
    End Function

    Private Sub GetConfig()
        Dim strConfig As String = ReadRegKeyStr(TrelloRegKey, "config")
        If Not String.IsNullOrEmpty(strConfig) Then
            config = JsonConvert.DeserializeObject(strConfig)
        End If
    End Sub

    Private Sub ShowAuthorisePrompt()
        RefreshToken()
        If String.IsNullOrEmpty(Me.token) Then
            If authOpen = True Then
                container.Show()
                container.BringToFront()
            Else
                container.Width = 550
                container.Height = 610
                container.Text = "Authorise Trello"

                wb.Visible = True
                wb.Width = container.Width - 15
                wb.Height = container.Height
                wb.ScrollBarsEnabled = False

                NullBrowser()
                AttachContainerHandlers()

                container.Controls.Add(wb)

                wb.Navigate(TrelloBase & TrelloVersion _
                            & "/authorize?key=" & TrelloKey _
                            & "&name=" & TrelloName _
                            & "&expiration=never&response_type=token&scope=read,write")

                container.Show()
                container.BringToFront()
                authOpen = True
            End If
        Else
            ShowError("Already Authorised!", "Trello Authorisation")
        End If
    End Sub

    Private Sub GetTokenFromWebBrowser()
        If (String.IsNullOrEmpty(token)) Then
            Try
                Cursor.Current = Cursors.AppStarting
                Dim preElements As HtmlElementCollection
                preElements = wb.Document.GetElementsByTagName("PRE")
                If preElements.Count > 0 Then
                    token = Trim(preElements(0).InnerText)
                Else
                    token = ""
                End If
                'System.Windows.Forms.MessageBox.Show("Token: " & token)
                container.Close()
                If ((Not String.IsNullOrEmpty(token)) AndAlso (token.Length = 64)) Then
                    WriteRegKey(TrelloRegKey, "token", token)
                    If Not String.IsNullOrEmpty(token) Then
                        System.Threading.SynchronizationContext.SetSynchronizationContext(New WindowsFormsSynchronizationContext())
                        Using worker As New System.ComponentModel.BackgroundWorker
                            AddHandler worker.DoWork, AddressOf FetchMemberInBackground
                            'AddHandler worker.RunWorkerCompleted, AddressOf GotMember
                            worker.RunWorkerAsync()
                        End Using
                        Dim root As String = ""
                        Dim key As String = ""
                        SplitSubKey(TrelloMappingRegKey, root, key)
                        DeleteSubKey(root, key)
                        CreateSubKey(root, key)
                        ShowInformation("Authorised!", "Trello Authorisation")
                    End If
                Else
                    token = ""
                    ShowError("Failed to Authorise! Please try again", "Trello Authorisation")
                End If
            Finally
                Cursor.Current = Cursors.Default
            End Try
        Else
            ShowError("Already Authorised!", "Trello Authorisation")
        End If
    End Sub

    Private Sub FetchMemberInBackground(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        Dim JSON As New JObject
        CallTrello("members/me", "", JSON)
        WriteRegKey(TrelloRegKey, "id", JSON("id").ToString)
    End Sub

    Private Function MemberListIncludesMe(ByRef membersList As JArray)
        Dim result As Boolean = False
        Dim myId As String = ReadRegKeyStr(TrelloRegKey, "id")
        If Not String.IsNullOrEmpty(myId) Then
            For Each memberId In membersList
                result = memberId.ToString = myId
                If result Then
                    Exit For
                End If
            Next
        End If
        Return result
    End Function

    Private Sub CreateTrelloCalendar()
        Dim objNS As Outlook._NameSpace = explorer.Session
        Dim objFolder As Outlook.MAPIFolder = objNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
        Dim bExists As Boolean = False

        GetTrelloCalendar()

        'Create Trello Calendar if it doesn't exist
        If objCalendar Is Nothing Then
            Dim root As String = ""
            Dim key As String = ""
            SplitSubKey(TrelloMappingRegKey, root, key)
            DeleteSubKey(root, key)
            DeleteRegKey(TrelloRegKey, "config")
            GetConfig()
            objCalendar = objFolder.Folders.Add(CalendarName, Outlook.OlDefaultFolders.olFolderCalendar)
        End If
        RefreshCalendarView()
    End Sub

    Private Sub GetTrelloCalendar()
        Dim objNS As Outlook._NameSpace = explorer.Session
        Dim objFolder As Outlook.MAPIFolder = objNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)

        objCalendar = Nothing

        'Look for Trello Calendar
        For i As Integer = 1 To objFolder.Folders.Count
            If (objFolder.Folders.Item(i).Name = CalendarName) Then
                objCalendar = objFolder.Folders.Item(i)
                Exit For
            End If
        Next
        RefreshCalendarView()
    End Sub

    Private Sub RefreshCalendarView()
        'Don't this this works...
        If objCalendar IsNot Nothing Then
            objCalendar.CurrentView.Apply()
        End If
    End Sub

    Private Sub CreateStaticConfig()
        'Setup container
        If container.InvokeRequired Then
            Dim rerun = New RerunWorker(AddressOf CreateStaticConfig)
            container.Invoke(rerun)
        Else
            container.Controls.Remove(wb)
            wb.Visible = False
            container.Text = "Configuration"
            container.Width = 500
            container.Height = 500

            'Add Configuration label
            Dim lbHeading As New Label
            lbHeading = CreateLabel("lbHeading", "Configuration", New System.Drawing.Point(10, 10),
                                                 True, 0, New System.Drawing.Font(lbHeading.Font.FontFamily, 16, Drawing.FontStyle.Bold),
                                                 container)
            'Center Configuration label
            lbHeading.Left = ((container.Width / 2) - lbHeading.Width / 2)

            'Add Boards label
            Dim lbBoards As New Label
            lbBoards = CreateLabel("lbBoards", "Boards", New System.Drawing.Point(10, 55),
                                                 True, 0, New System.Drawing.Font(lbBoards.Font.FontFamily, 10, Drawing.FontStyle.Bold),
                                                 container)

            'Create Options label
            Dim lbOptions As New Label
            lbOptions = CreateLabel("lbOptions", "Options", New System.Drawing.Point(cbWidth + (cbMarginLeft * 2), 55),
                                                 True, 0, New System.Drawing.Font(lbBoards.Font.FontFamily, 10, Drawing.FontStyle.Bold),
                                                 container)
        End If

    End Sub

    Private Sub CreateConfig()
        If CheckEnvironment(True, False, True, True) Then
            GetConfig()
            If configOpen = True Then
                'Bring to front if already open
                container.BringToFront()
            Else
                Try
                    'Create configuration screen
                    Cursor.Current = Cursors.AppStarting

                    'Create static page
                    System.Threading.SynchronizationContext.SetSynchronizationContext(New WindowsFormsSynchronizationContext())
                    Using worker As New System.ComponentModel.BackgroundWorker
                        AddHandler worker.DoWork, AddressOf CreateStaticConfig
                        'AddHandler worker.RunWorkerCompleted, AddressOf GotMember
                        worker.RunWorkerAsync()
                    End Using

                    'Get my boards

                    Dim JSON As New JObject
                    CallTrello("members/me", "", JSON)
                    If JSON("idBoards").Count > 0 Then
                        Dim pb As New TProgressBar
                        pb.Init("Loading Configuration", JSON("idBoards").Count + 1)
                        Dim boardJSON As New JObject
                        Dim cbCount As Integer = 1
                        Dim cbTabIndex As Integer = 1
                        Dim cb As New CheckBox
                        Dim maxHeight As Integer = 0
                        Dim checked As Boolean = False

                        'Loop all boards
                        For iBoard As Integer = 0 To JSON("idBoards").Count - 1
                            boardJSON.RemoveAll()
                            'Get board information
                            CallTrello("boards/" & JSON("idBoards")(iBoard).ToString, "", boardJSON)
                            If boardJSON.Count > 0 Then
                                'Create checkbox for board

                                If (config.HasValues AndAlso config("boards") IsNot Nothing) Then
                                    For Each configBoard As JObject In config("boards")
                                        checked = configBoard("id").ToString = boardJSON("id").ToString
                                        If checked Then
                                            Exit For
                                        End If
                                    Next

                                End If

                                cb = CreateCheckbox("cb-" & JSON("idBoards")(iBoard).ToString, boardJSON("name").ToString,
                                                    New System.Drawing.Point(10 + 10, cbStart + (cbHeight * cbCount) + ((cbCount - 1) * cbMarginTop)),
                                                    New System.Drawing.Size(cbWidth, cbHeight), cbTabIndex, True, container, checked)
                                cbCount = cbCount + 1
                                cbTabIndex = cbTabIndex + 1
                            End If
                            If iBoard = JSON("idBoards").Count - 1 Then
                                'Find maximum height of container, used to set container height later
                                maxHeight = cb.Bottom
                            End If
                            pb.StepIt()
                        Next

                        ''Create Options label
                        'Dim lbOptions As New Label
                        'lbOptions = CreateLabel("lbOptions", "Options", New System.Drawing.Point(cbWidth + (cbMarginLeft * 2), 55),
                        '                                     True, 3 + cbCount, New System.Drawing.Font(lbBoards.Font.FontFamily, 10, Drawing.FontStyle.Bold),
                        '                                     container)

                        cbCount = 1

                        'Create Push to Trello Checkbox
                        checked = (config.HasValues AndAlso config("push") IsNot Nothing AndAlso config("push").ToString.ToUpper = "TRUE")
                        cb = CreateCheckbox("cb-PushTrello", "Push to Trello",
                                            New System.Drawing.Point(cbWidth + (cbMarginLeft * 2) + 10, cbStart + (cbHeight * cbCount) + ((cbCount - 1) * cbMarginTop)),
                                            New System.Drawing.Size(cbWidth, cbHeight), cbTabIndex, True, container, checked)
                        cbCount = cbCount + 1
                        cbTabIndex = cbTabIndex + 1

                        checked = (config.HasValues AndAlso config("assigned") IsNot Nothing AndAlso config("assigned").ToString.ToUpper = "TRUE")
                        'Create Assigned to Me Checkbox
                        cb = CreateCheckbox("cb-AssignedToMe", "Cards are assigned to Me",
                                            New System.Drawing.Point(cbWidth + (cbMarginLeft * 2) + 10, cbStart + (cbHeight * cbCount) + ((cbCount - 1) * cbMarginTop)),
                                            New System.Drawing.Size(cbWidth, cbHeight), cbTabIndex, True, container)
                        cbCount = cbCount + 1
                        cbTabIndex = cbTabIndex + 1

                        'Create OK Button
                        Dim btnOK As New Button
                        With btnOK
                            .Text = "OK"
                            .Name = "btnOK"
                            .Parent = container
                            .Height = 23
                            .TabStop = cbTabIndex
                            .Location = New System.Drawing.Point((container.Width / 2) - .Width - 5, maxHeight)
                        End With
                        AddHandler btnOK.Click, AddressOf SaveConfig

                        pb.StepIt()

                        'Create Cancel Button
                        Dim btnCancel As New Button
                        With btnCancel
                            .Text = "Cancel"
                            .Name = "btnCancel"
                            .Parent = container
                            .Height = 23
                            .TabStop = btnOK.TabStop + 1
                            .Location = New System.Drawing.Point(btnOK.Right + 10, btnOK.Top)
                        End With
                        AddHandler btnCancel.Click, AddressOf ButtonCancelClose

                        'Finish setting up container (inner height)
                        Dim cs As Size
                        cs.Height = maxHeight + 33
                        cs.Width = 500
                        container.ClientSize = cs
                        AddHandler container.FormClosing, AddressOf CancelClose

                        pb.Done()
                        pb = Nothing

                        'Show the form
                        container.Show()
                        container.BringToFront()
                        configOpen = True
                    Else
                        ShowError("No boards to load", "Configuration")
                    End If
                Finally
                Cursor.Current = Cursors.Default
            End Try
            End If
        End If
    End Sub

    Private Sub CreateRequiredRegistryKeys()
        Dim Root As String = ""
        Dim Key As String = ""
        SplitSubKey(TrelloRegKey, Root, Key)
        CreateSubKey(Root, Key)
        SplitSubKey(TrelloMappingRegKey, Root, Key)
        CreateSubKey(Root, Key)
        WriteRegKey(TrelloRegKey, "author", "Bradly Sharpe")
        WriteRegKey(TrelloRegKey, "version", "1.0")
    End Sub

    Private Sub AttachContainerHandlers()
        AddHandler wb.Navigated, AddressOf LocationChanged
        AddHandler wb.Navigating, AddressOf LocationChanged
        AddHandler container.FormClosing, AddressOf CancelClose
    End Sub

    Private Sub AttachAppointmentHandlers()
        AddHandler objCalendar.Items.ItemAdd, AddressOf ItemAdded
        AddHandler objCalendar.Items.ItemChange, AddressOf ItemChanged

        For Each item As Outlook.AppointmentItem In objCalendar.Items
            AddHandler item.BeforeDelete, AddressOf ItemRemove
        Next
    End Sub

    Private Sub DeAuthoriseApplication(Optional ByVal ShowMessage As Boolean = True)
        RefreshToken()
        If Not String.IsNullOrEmpty(token) Then
            CheckEnvironment(True)
            DeleteRegKey(TrelloRegKey, "token")
            DeleteRegKey(TrelloRegKey, "id")
            DeleteRegKey(TrelloRegKey, "last_sync")
            DeleteRegKey(TrelloRegKey, "config")
            GetConfig()
            For Each folder As Outlook.Folder In explorer.Session.Folders
                If folder.DefaultItemType = Outlook.OlItemType.olAppointmentItem AndAlso folder.EntryID = objCalendar.EntryID Then
                    folder.Delete()
                End If
            Next
            token = ""
            If ShowMessage Then
                ShowInformation("Trello Outlook Addon has been deauthorised!", "Deauthorised")
            End If
        Else
            If ShowMessage Then
                ShowError("Trello Outlook Addon has not been authorised!", "Deauthorised")
            End If
        End If
    End Sub

    Private Sub SyncFromTrello(Optional ShowMessages As Boolean = True)
        If CheckEnvironment(True, True, True, True) Then
            If ((Not config.HasValues) Or (config("boards") Is Nothing) Or (Not config("boards").HasValues)) Then
                If ShowMessages Then
                    ShowError("No boards to sync, please configure!", "Sync Error")
                    btnConfig_Click(Me)
                End If
            Else
                Dim result As DialogResult = DialogResult.Yes
                If ShowMessages Then
                    result = ShowConfirmation("This will refresh from Trello and replace any changes." & vbCrLf & "Are you sure you want to sync?",
                                 "Sync Confirmation", MessageBoxButtons.YesNo)
                End If
                If result = DialogResult.Yes Then
                    Try
                        Cursor.Current = Cursors.AppStarting
                        'FindTrelloCalendar()
                        Dim JSONBoardCards As New JArray
                        Dim objOutlook As Outlook._Application = New Outlook.Application
                        Dim AssignedToMe As Boolean = False
                        If config.HasValues AndAlso config("assigned") IsNot Nothing AndAlso config("assigned").ToString().ToLower = "true" Then
                            AssignedToMe = True
                        End If
                        Dim GetAttachments As Boolean = True
                        If config.HasValues AndAlso config("attachments") IsNot Nothing AndAlso config("attachments").ToString().ToLower = "true" Then
                            GetAttachments = True
                        End If
                        For Each board In config("boards")
                            CallTrello("board/" & board("id").ToString & "/cards", "", JSONBoardCards)
                            Dim pb As New TProgressBar
                            pb.Init("Loading Board: " & board("name").ToString, JSONBoardCards.Count)
                            For Each item In JSONBoardCards
                                Dim card As New JObject
                                CallTrello("card/" + item("id").ToString, "", card)
                                CreateOutlookAppointment(board, card, True, AssignedToMe, GetAttachments)
                                pb.StepIt()
                            Next
                            pb.Done()
                            pb = Nothing
                        Next
                        WriteRegKey(TrelloRegKey, "last_sync", Now)
                    Finally
                        Cursor.Current = Cursors.Default
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub SetAppointmentAttributes(ByRef appt As Outlook.AppointmentItem, ByVal Board As JObject, ByVal Card As JObject, GetAttachments As Boolean, Optional ByVal UpdateReminder As Boolean = True)
        appt.Subject = Board("name").ToString() & " - " & Card("name").ToString()
        appt.Body = "Card Description: " & vbCrLf & Card("desc").ToString()

        If GetAttachments Then
            Dim JSONAttachments As New JArray
            CallTrello("cards/" & Card("id").ToString() & "/attachments", "", JSONAttachments)
            If JSONAttachments.Count > 0 Then
                appt.Body = appt.Body & vbCrLf & vbCrLf & "Attachments:"
                For Each attachment In JSONAttachments
                    appt.Body = appt.Body & vbCrLf & attachment("url").ToString()
                Next
            End If
        End If

        appt.Body = appt.Body & vbCrLf & vbCrLf & "Card Link: " & Card("shortUrl").ToString
        appt.Location = "Trello - " & Board("name").ToString()
        appt.StartUTC = Card("due").ToString()
        appt.Duration = 60
        If UpdateReminder Then
            appt.ReminderSet = True
        End If
        appt.ReminderMinutesBeforeStart = 15
        appt.BusyStatus = Outlook.OlBusyStatus.olBusy
    End Sub

    Private Function CreateOutlookAppointment(ByVal Board As JObject, ByVal Card As JObject, Optional HasDueDate As Boolean = True, Optional AssignedToMe As Boolean = True, Optional GetAttachments As Boolean = False)
        If Not (HasDueDate) Or (HasDueDate AndAlso Not String.IsNullOrEmpty(Card("due").ToString)) Then
            'have a card with due date!
            If ((Not AssignedToMe) Or (AssignedToMe AndAlso (MemberListIncludesMe(Card("idMembers"))))) Then
                '    If config("assiged") Is Nothing Or CheckAssignedMembers(JSONCard("idMembers")) Then
                'Dim appt As Outlook.AppointmentItem = objOutlook.CreateItem(Outlook.OlItemType.olAppointmentItem)
                If Not UpdateAppointment(Board, Card, GetAttachments) Then
                    Dim appt As Outlook.AppointmentItem
                    appt = objCalendar.Items.Add(Outlook.OlItemType.olAppointmentItem)
                    SetAppointmentAttributes(appt, Board, Card, GetAttachments)
                    appt.Save()

                    Card.Add(New JProperty("globalid", appt.GlobalAppointmentID))
                    WriteMapping(appt.GlobalAppointmentID, Card.ToString())
                    Return appt
                End If
            End If
        End If
        Return Nothing
    End Function

    Private Function UpdateAppointment(ByVal Board As JObject, ByVal Card As JObject, GetAttachments As Boolean)
        Dim Result As Boolean = False
        Dim ApptJSON As JObject = FindAppointmentByTrelloId(Card("id").ToString)
        If ApptJSON.HasValues Then
            Dim Appt As Outlook.AppointmentItem = FindAppointmentByGlobalId(ApptJSON("globalid").ToString())
            If Appt IsNot Nothing Then
                'ShowInformation("Should update this one!", "Update Required")
                SetAppointmentAttributes(Appt, Board, Card, GetAttachments, False)
                Appt.Save()
                Result = True
            End If
        End If
        Return Result
    End Function

    Private Function FindAppointmentByTrelloId(ByVal TrelloId As String) As JObject
        Dim JSON As JObject = FindRegKeyByTrelloId(TrelloId)

        If JSON.HasValues Then
            Return JSON
        End If

        Return New JObject
    End Function

    Private Function FindAppointmentByGlobalId(ByVal GlobalId As String) As Outlook.AppointmentItem
        Dim appt As Outlook.AppointmentItem = Nothing

        appt = (From item As Outlook.AppointmentItem In objCalendar.Items Where item.GlobalAppointmentID = GlobalId).FirstOrDefault

        Return appt
    End Function

    Private Sub FixMe()
        Dim OKtoFix As Boolean = (ShowConfirmation("Are you sure you want to fix calendars?" & vbCrLf & "This cannot be undone!", "Fix", MessageBoxButtons.YesNo) = DialogResult.Yes)
        If OKtoFix Then
            Dim objNS As Outlook._NameSpace = explorer.Session
            Dim objFolder As Outlook.MAPIFolder = objNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)

            For i As Integer = 1 To objFolder.Folders.Count
                If ShowConfirmation("Delete items and calendar from '" & objFolder.Folders.Item(i).Name & "'?", "Calendar Fix", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    For Each appt As Outlook.AppointmentItem In objFolder.Folders.Item(i).Items
                        Try
                            appt.Delete()
                        Catch
                        End Try
                    Next
                    Try
                        objFolder.Folders.Item(i).Delete()
                    Catch
                        ShowInformation("Coultn't delete calendar: '" & objFolder.Folders.Item(i).Name & "'", "Error")
                    End Try
                End If
            Next
            objFolder.CurrentView.Apply()
            ShowInformation("Done!", "Calendar Fix")
        End If
    End Sub

#End Region



#Region "Event Callbacks"

    Private Sub ItemAdded(ByVal Item As Outlook.AppointmentItem)
        Dim a As Integer = 1
        AddHandler Item.BeforeDelete, AddressOf ItemRemove
    End Sub

    Private Sub ItemChanged(ByVal Item As Outlook.AppointmentItem)
        Dim a As Integer = 1
    End Sub

    Private Sub ItemRemove(ByVal Item As Outlook.AppointmentItem, ByRef Cancel As Boolean)
        Dim a As Integer = 1
    End Sub

    Private Sub ButtonCancelClose(sender As Object, e As System.EventArgs)
        CancelClose(sender, Nothing)
    End Sub

    Private Sub CancelClose(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
        Disposed(sender, Nothing)
        If Not e Is Nothing Then
            e.Cancel = True
        End If
        container.SuspendLayout()
        container.Controls.Clear()
        container.ResumeLayout()
        configOpen = False
        authOpen = False
    End Sub

    Private Sub LocationChanged(sender As Object, e As System.EventArgs)
        Dim url As System.Uri
        url = wb.Url

        If ((Not url Is Nothing) AndAlso
            (LCase(url.ToString) = TrelloBase & TrelloVersion & "/token/approve") AndAlso
            (String.IsNullOrEmpty(token)) AndAlso Not (wb.Document.Body.InnerHtml Is Nothing)) Then
            GetTokenFromWebBrowser()
        End If
    End Sub

    Private Sub Disposed(sender As Object, e As System.EventArgs)
        NullBrowser()
        If wb.Visible Then
            wb.Hide()
        End If
        If container.Visible Then
            container.Hide()
        End If
    End Sub

#End Region



#Region "Helpers"

    Private Function CheckEnvironment(ByVal RequireToken As Boolean,
                                 Optional ByVal Configuration As Boolean = False,
                                 Optional ByVal Calendar As Boolean = False, Optional ByVal CreateCalendar As Boolean = False) As Boolean
        explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        Dim result As Boolean = True
        Dim errormsg As String = ""
        If RequireToken Then
            If String.IsNullOrEmpty(token) Then
                RefreshToken()
            End If
            If String.IsNullOrEmpty(token) Then
                result = False
                errormsg = "Authorisation required!" & vbCrLf & "Please authorise Trello"
            End If
        End If

        If Configuration AndAlso result Then
            If Not config.HasValues Then
                GetConfig()
            End If
            If Not config.HasValues Then
                result = False
                errormsg = "Configuration required!" & vbCrLf & "Please run the configuration"
            End If
        End If

        If Calendar AndAlso result Then
            If objCalendar Is Nothing Then
                GetTrelloCalendar()
            End If
            If objCalendar Is Nothing AndAlso CreateCalendar Then
                CreateTrelloCalendar()
            ElseIf objCalendar Is Nothing Then
                result = False
                errormsg = "Could not find Trello Calendar!"
            End If
        End If

        If Not result Then
            ShowError(errormsg, "Environment Checking")
        End If
        Return result
    End Function

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i),
                              StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return ""
    End Function

    Private Function CreateCheckbox(ByVal Name As String, ByVal Text As String, ByVal Location As Point,
                                    ByVal Size As Size, ByVal TabIndex As Integer, ByVal Visible As Boolean,
                                    ByVal Parent As Control, Optional Checked As Boolean = False) As CheckBox
        Dim cb As CheckBox = New CheckBox
        With cb
            .Location = Location
            .Name = Name
            .Text = Text
            .Size = Size
            .TabIndex = TabIndex
            .Visible = Visible
            .Parent = Parent
            .Checked = Checked
            .Show()
        End With
        Parent.Controls.Add(cb)
        Return cb
    End Function

    Private Function CreateLabel(ByVal Name As String, ByVal Text As String, ByVal Location As Point,
                                 ByVal AutoSize As Boolean, ByVal TabIndex As Integer, ByVal Font As System.Drawing.Font,
                                 ByVal Parent As Control) As Label
        Dim lb As New Label
        With lb
            .Location = Location
            .Name = Name
            .Text = Text
            .AutoSize = AutoSize
            .Font = Font
            .TabIndex = TabIndex
            .Parent = Parent
            .Show()
        End With
        Return lb
    End Function

    Private Sub ShowError(ByVal Text As String, ByVal Caption As String)
        MessageBox.Show(Text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub ShowInformation(ByVal Text As String, ByVal Caption As String)
        MessageBox.Show(Text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Function ShowConfirmation(ByVal Text As String, ByVal Caption As String,
                                      ByVal Buttons As Windows.Forms.MessageBoxButtons) As Windows.Forms.DialogResult
        Return MessageBox.Show(Text, Caption, Buttons, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
    End Function

    Private Sub NullBrowser()
        wb.Navigate("about:blank")
        wb.Document.Write("<!DOCTYPE html><html><head></head><body></body></html>")
    End Sub

    Public ReadOnly Property Context As Object Implements Microsoft.Office.Core.IRibbonControl.Context
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Id As String Implements Microsoft.Office.Core.IRibbonControl.Id
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Tag As String Implements Microsoft.Office.Core.IRibbonControl.Tag
        Get
            Return Nothing
        End Get
    End Property

#End Region

End Class


#Region "External Classes"

Friend Class ImageConverter
    'Used when returning custom images for use in a ribbon
    Inherits System.Windows.Forms.AxHost
    Sub New()
        MyBase.New(Nothing)
    End Sub

    Public Shared Function Convert(ByVal image As System.Drawing.Image) As stdole.IPictureDisp
        Return AxHost.GetIPictureDispFromPicture(image)
    End Function
End Class

Friend Class TProgressBar
    Private _currentVal As Integer = 0
    Private _maxVal As Integer = 100
    Private _stepVal As Integer = 1
    Private _displayText As String = ""
    Private _showPercentage As Boolean = True

    Private _container As Form
    Private _progress As ProgressBar
    Private _label As Label

    Public Sub Init(ByVal Title As String, ByVal Max As Integer, Optional ByVal Current As Integer = 0, Optional ByVal showPercentage As Boolean = True, Optional ByVal Style As Forms.ProgressBarStyle = ProgressBarStyle.Continuous)
        If Style = ProgressBarStyle.Marquee Then
            _displayText = Title
            _maxVal = 1
            _currentVal = 0
            _showPercentage = False
        Else
            _displayText = Title
            _maxVal = Max
            _currentVal = Current
            _showPercentage = showPercentage
        End If
        Create(Style)
    End Sub

    Public Sub StepIt(Optional ByVal forward As Boolean = True)
        If Not _progress.Style = ProgressBarStyle.Marquee Then
            If forward Then
                _currentVal = _currentVal + 1
            Else
                _currentVal = _currentVal - 1
            End If
        End If
        SetProgress(_currentVal)
    End Sub

    Public Overloads Function Position() As Integer
        Return _currentVal
    End Function

    Public Overloads Function Position(value As Integer) As Integer
        If Not _progress.Style = ProgressBarStyle.Marquee Then
            Dim forward As Boolean = True
            If value < _currentVal Then
                forward = False
            End If
            While _currentVal <> value
                StepIt(forward)
            End While
        End If
        Return _currentVal
    End Function

    Public Sub Done()
        _container.Hide()
        _progress.Hide()
        _progress.Dispose()
        _container.Dispose()

        _currentVal = 0
        _maxVal = 0
        _stepVal = 1
        _displayText = ""
        _showPercentage = True
    End Sub

    Private Sub SetProgress(value As Integer)
        If Not _progress.Style = ProgressBarStyle.Marquee Then
            If Not _maxVal = 0 Then
                If _currentVal >= _maxVal Then
                    _progress.Value = _progress.Maximum
                Else
                    _progress.Value = (_currentVal / _maxVal) * 100
                End If
            End If
            If _showPercentage Then
                _label.Text = CStr(_progress.Value) & "%"
                PostionLabel()
            End If
        End If
        _progress.Update()
        _container.Update()
    End Sub

    Private Sub PostionLabel()
        _label.Left = (_container.ClientSize.Width - 10) - _label.Width
        _progress.Width = _label.Left - 20
    End Sub

    Private Sub Create(Style As Forms.ProgressBarStyle)
        _container = New Form
        _progress = New ProgressBar
        _label = New Label

        _container.Text = _displayText

        Dim ClientSize As New Size
        ClientSize.Height = 40
        ClientSize.Width = 500
        _container.ClientSize = ClientSize

        _progress.Parent = _container
        _progress.Location = New Point(10, 10)
        _progress.Width = ClientSize.Width - 20
        _progress.Height = ClientSize.Height - 20
        _progress.Style = Style
        _progress.Minimum = 0
        If Style = ProgressBarStyle.Marquee Then
            _progress.Maximum = _maxVal
        Else
            _progress.Maximum = 100
        End If
        _progress.Step = 1
        SetProgress(_currentVal)

        If _showPercentage Then
            _progress.Width = _progress.Width - 90
            _progress.Update()
            _label.Height = 24
            _label.AutoSize = True
            _label.Font = New System.Drawing.Font(_label.Font.FontFamily, 10, Drawing.FontStyle.Regular)
            _label.Location = New Point(_progress.Width + 10, _progress.Top + 2)
            _label.Parent = _container
            PostionLabel()
        End If
        _container.Show()
        _container.BringToFront()
    End Sub
End Class

#End Region

Public Delegate Sub RerunWorker()
