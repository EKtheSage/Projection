Imports Microsoft.Office.Interop.Excel
Imports ADODB
Imports SAS
Imports SASObjectManager

Public Class coverageList
    Private _covName As String
    Private _covNum As String

    Public Property coverageName As String
        Get
            Return _covName
        End Get
        Set(value As String)
            _covName = value
        End Set
    End Property
    Public Property coverageNum As String
        Get
            Return _covNum
        End Get
        Set(value As String)
            _covNum = value
        End Set
    End Property
End Class

Public Class lobList
    Private _lobName As String
    Private _lobNum As String

    Public Property lobName As String
        Get
            Return _lobName
        End Get
        Set(value As String)
            _lobName = value
        End Set
    End Property
    Public Property lobNum As String
        Get
            Return _lobNum
        End Get
        Set(value As String)
            _lobNum = value
        End Set
    End Property
End Class
Public Module PullData
    Public cov As New List(Of coverageList)
    Public lob As New List(Of lobList)
    Public covLookup As Range = wkstConstants.Range("lookUp_coverage")
    Public frmLogin As frmLogin = New frmLogin

    'add more coverage number here
    Private Function addCovToList() As IEnumerable(Of coverageList)
        Return New List(Of coverageList) From {
            New coverageList With {.coverageName = "BI", .coverageNum = "001"},
            New coverageList With {.coverageName = "PD", .coverageNum = "002"},
            New coverageList With {.coverageName = "COLL", .coverageNum = "003"},
            New coverageList With {.coverageName = "COMP", .coverageNum = "004"},
            New coverageList With {.coverageName = "UMC", .coverageNum = "UMC"},
            New coverageList With {.coverageName = "MP", .coverageNum = "007"},
            New coverageList With {.coverageName = "PIP", .coverageNum = "020"},
            New coverageList With {.coverageName = "XMC", .coverageNum = "XMC"}
        }
    End Function

    Private Function addLobToList() As IEnumerable(Of lobList)
        Return New List(Of lobList) From {
            New lobList With {.lobName = "Auto", .lobNum = "001"},
            New lobList With {.lobName = "All", .lobNum = "ALL"},
            New lobList With {.lobName = "Cycle", .lobNum = "002"},
            New lobList With {.lobName = "AIP", .lobNum = "005"}
        }
    End Function

    Public Sub getInitialTriangleList()
        'refresh the two pivot tables in the worksheet, then reset them to show all rows
        Dim triangleConn As OLEDBConnection = Application.ActiveWorkbook.Connections("spTriangleList").OLEDBConnection

        Dim bgQuery As Boolean = triangleConn.BackgroundQuery

        triangleConn.BackgroundQuery = False
        triangleConn.Refresh()
        triangleConn.BackgroundQuery = bgQuery

        For Each pvtTbl As PivotTable In CType(wkstControl.PivotTables, PivotTables)
            pvtTbl.ClearAllFilters()
        Next

    End Sub

    Public Sub getData()
        assignValueToControlSheet()
        getEPEE()
        getClsdAvg()
    End Sub
    Public Sub assignValueToControlSheet()
        Dim ptData As PivotTable = CType(wkstControl.PivotTables("PT_TriangleList1"), PivotTable)
        Dim ptATA As PivotTable = CType(wkstControl.PivotTables("PT_TriangleList2"), PivotTable)
        Dim rngData As Range = ptData.RowRange
        Dim rngATA As Range = ptATA.RowRange
        Dim covNum As String = CType(rngData.Offset(1, 0).Resize(1, 1).Value, String)
        Dim lobNum As String = CType(rngData.Offset(1, 4).Resize(1, 1).Value, String)

        'simple LINQ query on IEnumerables (fun fun fun)
        cov = CType(addCovToList(), List(Of coverageList))
        lob = CType(addLobToList(), List(Of lobList))

        Dim queryCov =
            From coverage In cov
            Where coverage.coverageNum = covNum
            Select coverage.coverageName

        Dim queryLob =
            From lines In lob
            Where lines.lobNum = lobNum
            Select lines.lobName

        If rngData.Rows.Count > 2 Or rngATA.Rows.Count > 2 Then
            MsgBox("You cannot have more than 1 row of selections!!")
            Exit Sub
        ElseIf rngData.Rows.Count = 1 Or rngATA.Rows.Count = 1 Then
            MsgBox("No such selection available!")
            Exit Sub
        Else
            wkstControl.Range("coverage").Value = queryCov.ToList(0).ToString
            wkstControl.Range("risk").Value = rngData.Offset(1, 1).Resize(1, 1).Value
            wkstControl.Range("company").Value = rngData.Offset(1, 2).Resize(1, 1).Value
            wkstControl.Range("state").Value = rngData.Offset(1, 3).Resize(1, 1).Value
            wkstControl.Range("lob").Value = queryLob.ToList(0).ToString
            If CType(rngData.Offset(1, 5).Resize(1, 1).Value, String) = "N" Then
                wkstControl.Range("include_ss").Value = "Net"
            Else
                wkstControl.Range("include_ss").Value = "Gross"
            End If
            wkstControl.Range("ATA_risk").Value = rngATA.Offset(1, 1).Resize(1, 1).Value
            wkstControl.Range("ATA_company").Value = rngATA.Offset(1, 2).Resize(1, 1).Value
            wkstControl.Range("ATA_state").Value = rngATA.Offset(1, 3).Resize(1, 1).Value
        End If
        projBase = CType(wkstControl.Range("proj_base").Value, String)
        evalGroup = CType(wkstControl.Range("eval_group").Value, String)

        getTrianglesFromSqlSvr()
    End Sub
    Public Sub getTrianglesFromSqlSvr()
        'Chr(39) is the character values for single quote, Chr(32) is the character values for space
        Dim snapDate As String = Chr(39) & CType(wkstControl.Range("CurrentEvalDate").Value, String) & Chr(39)
        Dim coverage As String = CType(wkstControl.Range("coverage").Value, String)
        Dim risk As String = ""
        Dim company As String = ""
        Dim state As String = ""
        Dim lobNum As String = Chr(39) & CType(CType(wkstControl.PivotTables("PT_TriangleList1"), PivotTable).RowRange.
                                Offset(1, 4).Resize(1, 1).Value, String) & Chr(39)
        Dim metric As String = ""
        Dim sqlString As String
        Dim oledbConn As OLEDBConnection

        'coverageNum is for EPEE query, where coverage field is numeric, not string
        cov = CType(addCovToList(), List(Of coverageList))

        Dim queryCov =
            From a In cov
            Where a.coverageName = coverage
            Select a.coverageNum

        For Each wkbkConn As WorkbookConnection In Application.ActiveWorkbook.Connections
            Select Case wkbkConn.Name
                Case "triangleCount", "trianglePaid", "triangleIncurred", "trianglePaidAlt", "triangleIncurredAlt"
                    Select Case wkbkConn.Name
                        Case "triangleCount"
                            metric = Chr(39) & "C" & Chr(39)
                        Case "trianglePaid", "trianglePaidAlt"
                            If includeSS = "Net" Then
                                metric = Chr(39) & "NP" & Chr(39)
                            Else
                                metric = Chr(39) & "GP" & Chr(39)
                            End If
                        Case "triangleIncurred", "triangleIncurredAlt"
                            If includeSS = "Net" Then
                                metric = Chr(39) & "NI" & Chr(39)
                            Else
                                metric = Chr(39) & "GI" & Chr(39)
                            End If
                    End Select

                    Select Case wkbkConn.Name
                        Case "triangleCount", "trianglePaid", "triangleIncurred"
                            risk = Chr(39) & CType(wkstControl.Range("risk").Value, String) & Chr(39)
                            company = Chr(39) & CType(wkstControl.Range("company").Value, String) & Chr(39)
                            state = Chr(39) & CType(wkstControl.Range("state").Value, String) & Chr(39)
                        Case "trianglePaidAlt", "triangleIncurredAlt"
                            risk = Chr(39) & CType(wkstControl.Range("ATA_risk").Value, String) & Chr(39)
                            company = Chr(39) & CType(wkstControl.Range("ATA_company").Value, String) & Chr(39)
                            state = Chr(39) & CType(wkstControl.Range("ATA_state").Value, String) & Chr(39)
                    End Select


                    sqlString = "proj.sptriangle @snap_date=" & snapDate & ", @company=" & company &
                           ", @LOB=" & lobNum & ", @state=" & state & ", @coverage=" & Chr(39) & coverage & Chr(39) &
                           ", @risk_segment=" & risk & ", @metric=" & metric
                    oledbConn = wkbkConn.OLEDBConnection
                    oledbConn.CommandText = sqlString
                    oledbConn.Refresh()
            End Select
        Next
    End Sub

    Public Sub getEPEE()
        Dim state As String
        Dim risk As String = ""
        Dim sqlString As String
        Dim oledbConn As OLEDBConnection
        Dim coverage As String = CType(wkstControl.Range("coverage").Value, String)

        Dim queryCov =
            From a In cov
            Where a.coverageName = coverage
            Select a.coverageNum

        Dim coverage2 As String = Convert.ToString(queryCov.ToList(0).ToString)

        If coverage2 = "XMC" Then
            coverage2 = "(036, 083)"
        ElseIf coverage2 = "UMC" Then
            coverage2 = "(005, 024, 073, 074)"
        Else
            coverage2 = "(" & coverage2 & ")"
        End If

        state = CType(wkstControl.Range("state").Value, String)

        If state = "CW" Then
            state = ""
        ElseIf state.Substring(0, 1) = "x" Then 'xNY, x4, etc -> could prove to be annoying down the road
            Dim state2 As String = state.Substring(1)
            If IsNumeric(state2) Then
                state = " And A.state NOT IN (" & Chr(39) & state2 & Chr(39) & ")"
            Else
                state = " And A.state NOT IN (" & Chr(39) & state2 & Chr(39) & ")"
            End If
        Else
            state = " And A.state =" & Chr(39) & state & Chr(39)
        End If

        risk = Chr(39) & CType(wkstControl.Range("risk").Value, String) & Chr(39)
        sqlString = "SELECT A.Date, sum(A.EP) As EP, sum(A.EE) / 365 as EE " &
                "From EPEE As A " &
                "Where A.Risk_Seg =" & risk & state & " And A.Coverage IN" &
                coverage2 & Chr(32) &
                "Group By A.Date " &
                "Order By A.Date"
        oledbConn = Application.ActiveWorkbook.Connections("EPEE").OLEDBConnection
        oledbConn.CommandText = sqlString
        oledbConn.Refresh()
    End Sub

    Public Sub getClsdAvg()
        Dim qT As QueryTable
        Dim urlString, risk, company, lob, state, covNum, coverage As String

        qT = wkstClsdAvg.QueryTables("ClosedAvgWebQuery")

        If CType(wkstControl.Range("risk").Value, String) = "ALL" Then
            risk = "&risksector=*"
        Else
            risk = "&risksector=" & CType(wkstControl.Range("risk").Value, String)
        End If

        If CType(wkstControl.Range("company").Value, String) = "ALL" Then
            company = "&company=**"
        ElseIf CType(wkstControl.Range("company").Value, String) = "GIComb" Then
            company = "&company=GICS"
        ElseIf CType(wkstControl.Range("company").Value, String) = "GCComb" Then
            company = "&company=GCCN"
        ElseIf CType(wkstControl.Range("company").Value, String) = "CMStd" Then
            company = "&company=CS"
        ElseIf CType(wkstControl.Range("company").Value, String) = "CMNon" Then
            company = "&company=CN"
        ElseIf CType(wkstControl.Range("company").Value, String) = "CM" Then
            company = "&company=CSCN"
        Else
            company = "&company=" & CType(wkstControl.Range("company").Value, String)
        End If

        If CType(wkstControl.Range("lob").Value, String) = "Auto" Then
            lob = "&lob=AU"
        ElseIf CType(wkstControl.Range("lob").Value, String) = "Cycle" Then
            lob = "&lob=CY"
        Else
            lob = "&lob=**"
        End If

        If CType(wkstControl.Range("state").Value, String) = "CW" Then
            state = "&state=**"
        ElseIf CType(wkstControl.Range("state").Value, String) = "xNY" Then
            state = "&state=**NY"
        ElseIf CType(wkstControl.Range("state").Value, String) = "x4" Then
            state = "&state=**NYFLNJMI"
        Else
            state = "&state=" & CType(wkstControl.Range("state").Value, String)
        End If

        covNum = CType(CType(wkstControl.PivotTables("PT_TriangleList1"), PivotTable) _
                        .RowRange.Offset(1, 0).Resize(1, 1).Value, String)
        If covNum = "UMC" Then
            coverage = "&coverage=005024073074"
        ElseIf covNum = "XMC" Then
            coverage = "&coverage=036083"
        Else
            coverage = "&coverage=" & covNum
        End If

        urlString = "http://sasinet.geico.net/sas-cgi/broker.exe?_service=kpprd&_debug=0&_program=pri_pgm.reportdriver1.sas&prog=rsvclosedaverage&report=view1" &
                "&years=*&ratestructure=*" & risk & company & lob & state & coverage & "&covstat=****&atlas=*"
        With qT
            .Connection = "URL;" & urlString
            .Refresh()
        End With

    End Sub
    Public Sub loadGUIBNRCountFormLogin()
        frmLogin.Show()
    End Sub

    Public Sub getGUIBNRCount()
        'need to get the last time's GU IBNR Count from SQL SVR.
        Dim objFactory As ObjectFactoryMulti2 = New ObjectFactoryMulti2()
        Dim objServerDef As ServerDef = New ServerDef()
        Dim objWorkspace As Workspace
        Dim objKeeper As ObjectKeeper = New ObjectKeeper()
        Dim objLibref As Libref
        Dim adoConn As Connection = New Connection
        Dim adoRS As Recordset = New Recordset
        Dim user, coverage, pw, risk, state, sqlString As String
        Dim tblIBNRCnt As ListObject = wkstIBNRCnt.ListObjects("tbl_IBNRCount")
        Dim covName As String = CType(wkstControl.Range("coverage").Value, String)

        user = frmLogin.txtBxUser.Text
        pw = frmLogin.txtBxPw.Text
        frmLogin.Hide()
        frmLogin.txtBxPw.Clear()

        cov = CType(addCovToList(), List(Of coverageList))
        Dim queryCov =
            From a In cov
            Where a.coverageName = covName
            Select a.coverageNum

        tblIBNRCnt.DataBodyRange.Offset(1, 0).ClearContents()
        tblIBNRCnt.Resize(wkstIBNRCnt.Range("A1:B2"))

        objServerDef.MachineDNSName = "edwkdp"
        objServerDef.Port = 8594
        objServerDef.Protocol = Protocols.ProtocolBridge
        objServerDef.ClassIdentifier = "440196d4-90f0-11d0-9f41-00a024bb830c"

        objWorkspace = CType(objFactory.CreateObjectByServer("SASApp", True, objServerDef, user, pw), Workspace)
        objKeeper.AddObject(1, "GUCount", objWorkspace) 'GUCount is the name of the object
        objLibref = objWorkspace.DataService.AssignLibref("GU", "", "/reservinganalysis/groundup/ibnr/", "")
        'open a connection
        adoConn.Open("Provider=sas.IOMProvider;SAS Workspace ID=" & objWorkspace.UniqueIdentifier)

        'sqlstring to pull from Control tab, add 21916 to SAS date to get correct date in Excel
        If CType(wkstControl.Range("state").Value, String) = "CW" Then
            state = ""
        Else
            state = "rtd_st_cd =" & Chr(39) & CType(wkstControl.Range("state").Value, String) & Chr(39) & " and "
        End If

        risk = "rsk_grp_cd =" & Chr(39) & CType(wkstControl.Range("risk").Value, String) & Chr(39) & " and "

        Dim coverage2 As String

        If queryCov.ToList(0).ToString = "UMC" Then
            coverage2 = "('005', '024', '073', '074')"
        ElseIf queryCov.ToList(0).ToString = "XMC" Then
            coverage2 = "('036', '083')"
        Else
            coverage2 = "(" & Chr(39) & queryCov.ToList(0).ToString & Chr(39) & ")"
        End If
        coverage = "ATA_s160_cvrg_cd IN " & coverage2 & Chr(32)
        sqlString = "select (snp_dt+21916) as date, sum(IBNR_Cnts) as IBNRCnt from GU.CNTS_RESPREAD_FINAL_ST where " &
            state & risk & coverage &
            "Group by date " &
            "Order by date "

        wkstIBNRCnt.Range("E1").Value = sqlString

        adoRS.Open(sqlString, adoConn)

        For i As Integer = 1 To adoRS.Fields.Count
            CType(wkstIBNRCnt.Cells(1, i), Range).Value = adoRS.Fields(i - 1).Name
        Next
        wkstIBNRCnt.Range("A2").CopyFromRecordset(adoRS)
        objKeeper.RemoveAllObjects()
    End Sub
End Module
