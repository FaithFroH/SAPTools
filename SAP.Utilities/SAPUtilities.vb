
Imports System.Data

Public Class SAPSessionStatusMessage
    Public Property MessageType As String
    Public Property MessageText As String
End Class

Public Class SAPUtilities

    Public Shared Function ReadSessionStatusPane(session) As SAPSessionStatusMessage
        Dim statusBar = session.FindById("wnd[0]/sbar")
        Return New SAPSessionStatusMessage() With {
            .MessageText = statusBar.Text,
            .MessageType = statusBar.MessageType
        }
    End Function

    Public Shared Function FindElementByID(session, id)
        Try
            Return session.FindById(id)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Wait Session transaction complete
    ''' </summary>
    ''' <param name="session"></param>
    Public Shared Sub WaitLoading(session)
        ' Wait loading
        Do While session.Busy
            Threading.Thread.Sleep(500)
        Loop
    End Sub

    Public Shared Sub DetectAndCloseAllExtraModal(session As Object)
        While session.Children.Count > 1
            ' Get last modal
            Dim wnd = session.Children(session.Children.Count - 1)
            wnd.Close()
            WaitLoading(session)
        End While
    End Sub

    ''' <summary>
    ''' Open New TCode
    ''' Make sure you're at SAP Home Page - SAP Easy access
    ''' </summary>
    ''' <param name="session"></param>
    ''' <param name="tcode"></param>
    Public Shared Sub OpenTcode(session As Object, tcode As String)
        'Input Code
        session.FindById("wnd[0]/tbar[0]/okcd").Text = tcode
        'Press Enter 
        session.FindById("wnd[0]").SendVKey(0)
        WaitLoading(session)
    End Sub
    ''' <summary>
    ''' Force Exit T-code, cancel all data
    ''' </summary>
    ''' <param name="session"></param>
    Public Shared Sub ExitTCode(session)
        Dim complete = False
        Do While Not complete
            Dim backButton = session.FindById("wnd[0]/tbar[0]/btn[3]")
            'Complete Exit when back button is disable
            If backButton.Changeable Then
                'Click Exit Button
                session.FindById("wnd[0]/tbar[0]/btn[15]").Press()
                WaitLoading(session)
                Dim dialog = FindElementByID(session, "wnd[1]")
                If Not IsNothing(dialog) Then
                    'Click No Button
                    session.FindById("wnd[1]/usr/btnSPOP-OPTION2").Press()
                End If
                WaitLoading(session)
            Else
                complete = True
            End If
        Loop

    End Sub
    ''' <summary>
    ''' Find MainAreaName in T-code that has multiple UserArea or UserArea change dynamically
    ''' </summary>
    ''' <param name="SAPSession"></param>
    ''' <param name="O_Log"></param>
    ''' <returns></returns>
    Public Shared Function FindMainAreaName(ByRef SAPSession As Object, ByRef O_Log As String) As String
        Dim MainArea As Object
        ' Main AreaName are different each time request
        ' Find MainAreaName
        Dim MainAreaName As String = ""
        Dim User = SAPSession.findById("wnd[0]/usr")
        For i = 0 To User.Children.Count - 1
            MainAreaName = User.Children(CInt(i)).Name
            If Left(MainAreaName, 15) = "SUB0:SAPLMEGUI:" Then
                Exit For
            End If
        Next
        O_Log += Environment.NewLine + "Main Area Name :" + MainAreaName

        ' Find MainArea
        MainArea = SAPSession.findById("wnd[0]/usr/sub" + MainAreaName)
        O_Log += Environment.NewLine + " Main Area Found :" + MainArea.Type
        Return MainAreaName
    End Function
    ''' <summary>
    ''' Scroll SAP Table Control
    ''' </summary>
    ''' <param name="Session"></param>
    ''' <param name="SAPTable"></param>
    ''' <param name="SAPTableID"></param>
    ''' <param name="NewScrollBarPosition"></param>
    ''' <param name="O_Log"></param>
    ''' <param name="ScrollWaitingTime"></param>
    Public Shared Sub ScrollSAPTableVertically(ByRef Session As Object, ByRef SAPTable As Object, SAPTableID As Object,
                                               NewScrollBarPosition As Long, ByRef O_Log As String, ScrollWaitingTime As Integer)
        Dim currentPosition = SAPTable.VerticalScrollbar.Position
        Dim retryCount = 0
        If NewScrollBarPosition > SAPTable.VerticalScrollbar.Maximum Then
            NewScrollBarPosition = SAPTable.VerticalScrollbar.Maximum
        End If

        If NewScrollBarPosition < SAPTable.VerticalScrollbar.Minimum Then
            NewScrollBarPosition = SAPTable.VerticalScrollbar.Minimum
        End If
        O_Log += Environment.NewLine + "NewScrollBarPosition :" + CStr(NewScrollBarPosition)

        ' When All visible row is processed . Scroll down 
        Do While currentPosition = SAPTable.VerticalScrollbar.Position And retryCount < 3
            'Scroll Table
            SAPTable.VerticalScrollbar.Position = NewScrollBarPosition
            ' Sleep to wait Scrollbar action complete
            O_Log += Environment.NewLine + "Start Sleep :" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")
            WaitLoading(Session)
            O_Log += Environment.NewLine + "End Sleep :" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")
            'Reload After Scroll 
            SAPTable = Session.findById(SAPTableID)
            ' Detect Error or Warning while scroll. 
            Dim statusBar = Session.FindById("wnd[0]/sbar")
            If Not String.IsNullOrEmpty(statusBar.Text) Then
                If (statusBar.MessageType = "W" Or statusBar.MessageType = "E") Then
                    'Reset Position to scroll again
                    SAPTable.VerticalScrollbar.Position = currentPosition
                    SAPTable = Session.findById(SAPTableID)
                End If
            End If
            retryCount += 1
        Loop

        O_Log += Environment.NewLine + "Vertical Scroll Bar Position :" + CStr(SAPTable.VerticalScrollbar.Position)
    End Sub
    ''' <summary>
    ''' Read Sap Table Control Data To DataTable. 
    ''' Ignore error or disabled cell
    ''' Read as string only
    ''' </summary>
    ''' <param name="O_Log"></param>
    ''' <param name="SAPSession"></param>
    ''' <param name="SAPTableID"></param>
    ''' <param name="KeyColumnIndex"></param>
    ''' <param name="UniqueKey"></param>
    ''' <param name="ScrollWaitingTime"></param>
    ''' <returns></returns>
    Public Shared Function SAPTableControlToDataTable(ByRef O_Log As String,
                                                      ByRef SAPSession As Object, SAPTableID As String,
                                                      KeyColumnIndex As Integer, Optional UniqueKey As Boolean = False,
                                                      Optional ScrollWaitingTime As Integer = 2000) As DataTable
        Dim Table = New DataTable()

        Dim SAPTable As Object
        Dim SAPTableColsCollection As Object
        Dim SAPTableVisibleRowCount As Object
        Dim Keys = New List(Of String)

        ' Try Get Table
        Try
            SAPTable = SAPSession.findById(SAPTableID)
        Catch ex As Exception
            O_Log += Environment.NewLine + " Table Not Found !!!!"
            Return Table
        End Try

        ' Try Parse Table
        Try

            SAPTableColsCollection = SAPTable.Columns
            SAPTableVisibleRowCount = SAPTable.VisibleRowCount

            Dim O_RowCount = SAPTable.RowCount
            Dim O_ColCount = SAPTableColsCollection.Count

            O_Log += Environment.NewLine + " SapTable Found !!! Row Count : " + CStr(O_RowCount) + ". Col Count: " + CStr(O_ColCount) +
         Environment.NewLine + " Visibile Row Count " + CStr(SAPTableVisibleRowCount) +
         Environment.NewLine + " Vertical Scroll Bar - Max : " + CStr(SAPTable.VerticalScrollbar.Maximum) + ". Min :" + CStr(SAPTable.VerticalScrollbar.Minimum) + ". Position : " + CStr(SAPTable.VerticalScrollbar.Position)

            ' Read SapTable Columns
            SAPTableColsCollection = SAPTable.Columns
            For index As Integer = 0 To O_ColCount - 1
                Dim Col = SAPTableColsCollection.ElementAt(index)
                Dim Title = Col.Title + "_" + CStr(index)
                Table.Columns.Add(Title)
            Next



            ' Parse SapTable 
            ' Set Vertical to Minimum if need to parse data from start
            If SAPTable.VerticalScrollbar.Position <> SAPTable.VerticalScrollbar.Minimum Then
                ScrollSAPTableVertically(SAPSession, SAPTable, SAPTableID, SAPTable.VerticalScrollbar.Minimum, O_Log, ScrollWaitingTime)
            End If

            Dim NewScrollBarPosition As Long = 0
            Dim ProcessedRow As Object = 0
            Dim BlankRowFlag As Object = False
            Dim DuplicationKeyFlag As Object = False

            ' While havent processed all rows and blank row is not found
            While (ProcessedRow < O_RowCount And Not BlankRowFlag And Not DuplicationKeyFlag)
                'Minus 1 since index start from 0
                Dim LastDisplayRow = SAPTableVisibleRowCount - 1 + NewScrollBarPosition
                O_Log += Environment.NewLine + " Row From: " + CStr(NewScrollBarPosition) + " To: " + CStr(LastDisplayRow)
                ' Parse Row
                For rIndex As Integer = 0 To SAPTableVisibleRowCount - 1
                    O_Log += Environment.NewLine + "Process Row : " + CStr(ProcessedRow)

                    ' Check if ID col Cell is blank then Stop Read
                    Dim KeyCell = SAPTable.GetCell(rIndex, KeyColumnIndex)
                    If Len(KeyCell.Text) = 0 Then
                        BlankRowFlag = True
                        Exit For
                    End If
                    O_Log += Environment.NewLine + "BlankRowFlag Success !"

                    ' Check if current row key is exists
                    If (UniqueKey) Then
                        DuplicationKeyFlag = Keys.Any(Function(x As String)
                                                          Return x = KeyCell.Text
                                                      End Function)
                        If DuplicationKeyFlag Then
                            Exit For
                        End If
                        O_Log += Environment.NewLine + "DuplicationKeyFlag Success !"
                    End If
                    ' Add Keys
                    Keys.Add(KeyCell.Text)

                    ' All Conditions are pass then start parse column
                    Dim newRow = Table.NewRow()
                    ' Parse Column
                    For cIndex As Integer = 0 To O_ColCount - 1
                        Try
                            Dim Cell = SAPTable.GetCell(rIndex, cIndex)
                            newRow(cIndex) = Cell.Text
                        Catch ex As Exception
                            O_Log += Environment.NewLine + "Row :" + CStr(rIndex) + ". Cell : " + CStr(cIndex) + "|" + ex.Message
                        End Try
                    Next
                    ' Add new Row to Table
                    Table.Rows.Add(newRow)

                    ' Update Processed Row Flag
                    ProcessedRow += 1
                Next
                If (Not BlankRowFlag And Not DuplicationKeyFlag) Then
                    NewScrollBarPosition = CLng(SAPTable.VerticalScrollbar.Position) + CLng(SAPTableVisibleRowCount)
                    ScrollSAPTableVertically(SAPSession, SAPTable, SAPTableID, NewScrollBarPosition, O_Log, ScrollWaitingTime)
                End If
            End While

        Catch e As Exception
            O_Log += e.Message
        End Try

        Return Table
    End Function

    Public Shared Sub ScrollSAPTableVertically(Session As Object, ByRef SAPTable As Object, SAPTableID As String,
                                               NewScrollBarPosition As Long)
        Dim currentPosition = SAPTable.VerticalScrollbar.Position
        Dim retryCount = 0
        If NewScrollBarPosition > SAPTable.VerticalScrollbar.Maximum Then
            NewScrollBarPosition = SAPTable.VerticalScrollbar.Maximum
        End If

        If NewScrollBarPosition < SAPTable.VerticalScrollbar.Minimum Then
            NewScrollBarPosition = SAPTable.VerticalScrollbar.Minimum
        End If
        ' When All visible row is processed . Scroll down 
        Do While currentPosition = SAPTable.VerticalScrollbar.Position And retryCount < 3
            'Scroll Table
            SAPTable.VerticalScrollbar.Position = NewScrollBarPosition
            ' Sleep to wait Scrollbar action complete
            WaitLoading(Session)
            'Reload After Scroll 
            SAPTable = Session.findById(SAPTableID)
            ' Detect Error or Warning while scroll. 
            Dim statusBar = Session.FindById("wnd[0]/sbar")
            If Not String.IsNullOrEmpty(statusBar.Text) Then
                If (statusBar.MessageType = "W" Or statusBar.MessageType = "E") Then
                    'Reset Position to scroll again
                    SAPTable.VerticalScrollbar.Position = currentPosition
                    SAPTable = Session.findById(SAPTableID)
                End If
            End If
            retryCount += 1
        Loop
    End Sub
    Public Shared Function SAPTableControlToDataTableV2(SAPSession As Object, SAPTableID As String,
                                                    KeyColumnIndex As Integer, Optional UniqueKey As Boolean = False) As DataTable
        Dim Table = New DataTable()

        Dim SAPTable As Object
        Dim SAPTableColsCollection As Object
        Dim SAPTableVisibleRowCount As Object
        Dim Keys = New List(Of String)

        ' Try Get Table
        SAPTable = SAPSession.FindById(SAPTableID)
        SAPTableColsCollection = SAPTable.Columns
        SAPTableVisibleRowCount = SAPTable.VisibleRowCount
        Dim RowCount = SAPTable.RowCount
        Dim ColCount = SAPTableColsCollection.Count
        ' Try Parse Table
        Try
            ' Read SapTable Columns
            SAPTableColsCollection = SAPTable.Columns
            For index As Integer = 0 To ColCount - 1
                Dim Col = SAPTableColsCollection.ElementAt(index)
                Dim Title = Col.Title + "_" + CStr(index)
                Table.Columns.Add(Title)
            Next



            ' Parse SapTable 
            ' Set Vertical to Minimum to start parsing from start
            If SAPTable.VerticalScrollbar.Position <> SAPTable.VerticalScrollbar.Minimum Then
                ScrollSAPTableVertically(SAPSession, SAPTable, SAPTableID, SAPTable.VerticalScrollbar.Minimum)
            End If

            Dim NewScrollBarPosition As Long = 0
            Dim ProcessedRow As Integer = 0
            Dim BlankRowFlag As Boolean = False
            Dim DuplicationKeyFlag As Boolean = False

            ' While havent processed all rows and blank row is not found
            While (ProcessedRow < RowCount And Not BlankRowFlag And Not DuplicationKeyFlag)
                'Minus 1 since index start from 0
                Dim LastDisplayRow = SAPTableVisibleRowCount - 1 + NewScrollBarPosition
                ' Parse Row
                For rIndex As Integer = 0 To SAPTableVisibleRowCount - 1
                    ' Check if ID col Cell is blank then Stop Read 
                    Dim KeyCell = SAPTable.GetCell(rIndex, KeyColumnIndex)
                    If Len(KeyCell.Text) = 0 Then
                        BlankRowFlag = True
                        Exit For
                    End If
                    ' Check if current row key is exists
                    If (UniqueKey) Then
                        DuplicationKeyFlag = Keys.Any(Function(x As String)
                                                          Return x = KeyCell.Text
                                                      End Function)
                        If DuplicationKeyFlag Then
                            Exit For
                        End If
                    End If
                    ' Add Keys
                    Keys.Add(KeyCell.Text)

                    ' All Conditions are pass then start parse column
                    Dim newRow = Table.NewRow()
                    ' Parse Column
                    For cIndex As Integer = 0 To ColCount - 1
                        Try
                            Dim Cell = SAPTable.GetCell(rIndex, cIndex)
                            Dim cellText = Cell.Text
                            If String.IsNullOrEmpty(cellText) Then
                                newRow(cIndex) = Cell.Tooltip
                            Else
                                newRow(cIndex) = cellText
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                    ' Add new Row to Table
                    Table.Rows.Add(newRow)

                    ' Update Processed Row Flag
                    ProcessedRow += 1
                Next
                If (Not BlankRowFlag And Not DuplicationKeyFlag) Then
                    NewScrollBarPosition = CLng(SAPTable.VerticalScrollbar.Position) + CLng(SAPTableVisibleRowCount)
                    ScrollSAPTableVertically(SAPSession, SAPTable, SAPTableID, NewScrollBarPosition)
                End If
            End While

        Catch e As Exception

        End Try

        Return Table
    End Function
End Class
