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
            New coverageList With {.coverageName = "XMC", .coverageNum = "XMC"},
            New coverageList With {.coverageName = "UIM", .coverageNum = "089"},
            New coverageList With {.coverageName = "UM", .coverageNum = "005"},
            New coverageList With {.coverageName = "T&L", .coverageNum = "006"},
            New coverageList With {.coverageName = "PDMD", .coverageNum = "008"},
            New coverageList With {.coverageName = "PE", .coverageNum = "009"},
            New coverageList With {.coverageName = "IPIP", .coverageNum = "010"},
            New coverageList With {.coverageName = "LBI", .coverageNum = "011"},
            New coverageList With {.coverageName = "NOBI", .coverageNum = "012"},
            New coverageList With {.coverageName = "UM EO", .coverageNum = "014"},
            New coverageList With {.coverageName = "UIM EO", .coverageNum = "015"},
            New coverageList With {.coverageName = "UM/UIM EO", .coverageNum = "016"},
            New coverageList With {.coverageName = "LPD", .coverageNum = "017"},
            New coverageList With {.coverageName = "LOI", .coverageNum = "018"},
            New coverageList With {.coverageName = "SBI", .coverageNum = "019"},
            New coverageList With {.coverageName = "UMPL", .coverageNum = "021"},
            New coverageList With {.coverageName = "BIPDSL", .coverageNum = "022"},
            New coverageList With {.coverageName = "BFC", .coverageNum = "023"},
            New coverageList With {.coverageName = "UM/UIM PL", .coverageNum = "024"},
            New coverageList With {.coverageName = "UIM PL", .coverageNum = "025"},
            New coverageList With {.coverageName = "CB", .coverageNum = "026"},
            New coverageList With {.coverageName = "UMSL", .coverageNum = "027"},
            New coverageList With {.coverageName = "DTOP", .coverageNum = "028"},
            New coverageList With {.coverageName = "LOU", .coverageNum = "029"},
            New coverageList With {.coverageName = "MBEN", .coverageNum = "030"},
            New coverageList With {.coverageName = "NOPD", .coverageNum = "031"},
            New coverageList With {.coverageName = "NOPIP", .coverageNum = "032"},
            New coverageList With {.coverageName = "NOCOLL", .coverageNum = "033"},
            New coverageList With {.coverageName = "NOCOMP", .coverageNum = "034"},
            New coverageList With {.coverageName = "UMPD RD", .coverageNum = "035"},
            New coverageList With {.coverageName = "XUIMPL", .coverageNum = "036"},
            New coverageList With {.coverageName = "DSCT", .coverageNum = "037"},
            New coverageList With {.coverageName = "EPIP", .coverageNum = "038"},
            New coverageList With {.coverageName = "PPI", .coverageNum = "039"},
            New coverageList With {.coverageName = "BFTL", .coverageNum = "040"},
            New coverageList With {.coverageName = "FUHP", .coverageNum = "041"},
            New coverageList With {.coverageName = "MOPP", .coverageNum = "042"},
            New coverageList With {.coverageName = "MOHP", .coverageNum = "043"},
            New coverageList With {.coverageName = "EC", .coverageNum = "044"},
            New coverageList With {.coverageName = "NEC", .coverageNum = "045"},
            New coverageList With {.coverageName = "FUNRL", .coverageNum = "046"},
            New coverageList With {.coverageName = "BFPP", .coverageNum = "047"},
            New coverageList With {.coverageName = "ID", .coverageNum = "049"},
            New coverageList With {.coverageName = "OBEL", .coverageNum = "050"},
            New coverageList With {.coverageName = "PPO", .coverageNum = "051"},
            New coverageList With {.coverageName = "XPPO", .coverageNum = "052"},
            New coverageList With {.coverageName = "CMPT", .coverageNum = "053"},
            New coverageList With {.coverageName = "CMPP", .coverageNum = "054"},
            New coverageList With {.coverageName = "OBELMCO", .coverageNum = "055"},
            New coverageList With {.coverageName = "COLT", .coverageNum = "057"},
            New coverageList With {.coverageName = "L/C", .coverageNum = "058"},
            New coverageList With {.coverageName = "PIPLR", .coverageNum = "059"},
            New coverageList With {.coverageName = "COLP", .coverageNum = "060"},
            New coverageList With {.coverageName = "REA", .coverageNum = "061"},
            New coverageList With {.coverageName = "LTCT", .coverageNum = "062"},
            New coverageList With {.coverageName = "LTCP", .coverageNum = "063"},
            New coverageList With {.coverageName = "RVL", .coverageNum = "065"},
            New coverageList With {.coverageName = "RPE", .coverageNum = "066"},
            New coverageList With {.coverageName = "RMD", .coverageNum = "067"},
            New coverageList With {.coverageName = "AD", .coverageNum = "068"},
            New coverageList With {.coverageName = "UIMPD", .coverageNum = "069"},
            New coverageList With {.coverageName = "UM GA", .coverageNum = "073"},
            New coverageList With {.coverageName = "UM RD", .coverageNum = "074"},
            New coverageList With {.coverageName = "UMPD GA", .coverageNum = "075"},
            New coverageList With {.coverageName = "UMPD", .coverageNum = "078"},
            New coverageList With {.coverageName = "RR", .coverageNum = "079"},
            New coverageList With {.coverageName = "MBI", .coverageNum = "082"},
            New coverageList With {.coverageName = "XUIM", .coverageNum = "083"},
            New coverageList With {.coverageName = "MPIF", .coverageNum = "085"},
            New coverageList With {.coverageName = "PPIF", .coverageNum = "086"},
            New coverageList With {.coverageName = "SCLS", .coverageNum = "087"},
            New coverageList With {.coverageName = "PIPA", .coverageNum = "PIPA"},
            New coverageList With {.coverageName = "BIA", .coverageNum = "BIA"}
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
        Dim triangleConn As OLEDBConnection = Application.ActiveWorkbook.Connections("TriangleList").OLEDBConnection
        Dim sqlString As String
        Dim snapDate As String = Chr(39) & CType(wkstControl.Range("CurrentEvalDate").Value, String) & Chr(39)
        Dim bgQuery As Boolean = triangleConn.BackgroundQuery


        sqlString = "proj.sptriangleList @snap_date=" & snapDate
        triangleConn.BackgroundQuery = False
        triangleConn.CommandText = sqlString
        triangleConn.Refresh()
        triangleConn.BackgroundQuery = bgQuery

        For Each pvtTbl As PivotTable In CType(wkstControl.PivotTables, PivotTables)
            pvtTbl.ClearAllFilters()
        Next
    End Sub

    Public Sub getData()
        assignValueToControlSheet()
    End Sub
    Public Sub assignValueToControlSheet()
        Dim ptData As PivotTable = CType(wkstControl.PivotTables("PT_TriangleList1"), PivotTable)
        Dim ptATA As PivotTable = CType(wkstControl.PivotTables("PT_TriangleList2"), PivotTable)
        Dim rngData As Range = ptData.RowRange
        Dim rngATA As Range = ptATA.RowRange
        Dim covNum As String = CType(rngData.Offset(1, 0).Resize(1, 1).Value, String)
        Dim lobNum As String = CType(rngData.Offset(1, 4).Resize(1, 1).Value, String)

        'When we click on Step2 button, we will make sure we're using the activework book's values
        wkstControl = CType(Application.ActiveWorkbook.Worksheets("Control"), Worksheet)
        wkstConstants = CType(Application.ActiveWorkbook.Worksheets("Constants"), Worksheet)
        wkstCount = CType(Application.ActiveWorkbook.Worksheets("Count"), Worksheet)
        wkstPaid = CType(Application.ActiveWorkbook.Worksheets("Paid"), Worksheet)
        wkstIncurred = CType(Application.ActiveWorkbook.Worksheets("Incurred"), Worksheet)
        wkstExpLoss = CType(Application.ActiveWorkbook.Worksheets("Exp Loss"), Worksheet)
        wkstSummary = CType(Application.ActiveWorkbook.Worksheets("Summary"), Worksheet)
        wkstQPage = CType(Application.ActiveWorkbook.Worksheets("Q Page"), Worksheet)
        wkstReviewTemplate = CType(Application.ActiveWorkbook.Worksheets("Review Template"), Worksheet)
        wkstData = CType(Application.ActiveWorkbook.Worksheets("Data"), Worksheet)
        wkstIBNRCnt = CType(Application.ActiveWorkbook.Worksheets("GU IBNR Count"), Worksheet)
        wkstClsdAvg = CType(Application.ActiveWorkbook.Worksheets("Closed Avg"), Worksheet)
        wkstClsMod = CType(Application.ActiveWorkbook.Worksheets("Closure Model"), Worksheet)
        'simple LINQ query on IEnumerables (fun fun fun)
        cov = CType(addCovToList(), List(Of coverageList))
        lob = CType(addLobToList(), List(Of lobList))

        'select the coverage name from the IEnumerables, the coverage name will be assigned
        'to worksheet Control namedrange "coverage"
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

        'Set calculation to automatic to make sure the variables below are updated
        Application.Calculation = XlCalculation.xlCalculationAutomatic

        projBase = CType(wkstControl.Range("proj_base").Value, String)
        evalGroup = CType(wkstControl.Range("eval_group").Value, String)
        includeSS = CType(wkstControl.Range("include_ss").Value, String)
        coverageField = CType(wkstControl.Range("coverage").Value, String)

        'also resets the default ata formula to this
        wkstControl.Range("Default_ATA").Formula = "=VLOOKUP(coverage,lookUp_coverage,5,0)"

        getTrianglesFromSqlSvr()
        getClsdAvg()
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
        ElseIf CType(wkstControl.Range("state").Value, String) = "x3" And
            CType(wkstControl.Range("coverage").Value, String) = "UIM" Then
            state = "&state=**KSNVCT"
        Else
            state = "&state=" & CType(wkstControl.Range("state").Value, String)
        End If

        covNum = CType(CType(wkstControl.PivotTables("PT_TriangleList1"), PivotTable) _
                        .RowRange.Offset(1, 0).Resize(1, 1).Value, String)


        Dim coverageName As String = CType(wkstControl.Range("coverage").Value, String)
        cov = CType(addCovToList(), List(Of coverageList))
        Dim queryCov =
            From coverages In cov
            Where coverages.coverageName = coverageName
            Select coverages.coverageNum

        covNum = queryCov.ToList(0).ToString
        If covNum = "UMC" Then
            coverage = "&coverage=005024073074"
        ElseIf covNum = "XMC" Then
            coverage = "&coverage=036083"
        ElseIf covNum = "PIPA" Then
            coverage = "&coverage=020086"
        ElseIf covNum = "BIA" Then
            coverage = "&coverage=001019"
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

    Public Sub runVBAHistory()
        Application.Run("History")
    End Sub

    Public Sub runVBANewData()
        Application.Run("NewData")
    End Sub

    Public Sub runVBAPrintPRP()
        Application.Run("printPRP")
    End Sub

    Public Sub runVBAPrintVI()
        Application.Run("printVI")
    End Sub

    Public Sub runVBAapprove()
        Application.Run("approve")
    End Sub

    Public Sub getGUIBNRCountVBA()
        Dim id As Integer
        id = CType(Application.Run("Info"), Integer)
        Application.Run("IBNRCounts", id)
    End Sub

    Public Sub getClsModVBA()
        Application.Run("getClsModData")
        Application.Run("getClsModSpr")
    End Sub

    Public Sub getGUIBNRCount()
        'they don't like the next big thing...:(
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
        'The line below can do windows authentication, but it doesn't seem to work.
        'objServerDef.BridgeSecurityPackage = "Negotiate"
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
        pw = ""
    End Sub
End Module
