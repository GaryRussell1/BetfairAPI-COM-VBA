' In Excel use Tools|References to add a reference to BA_COM_Betfair

Dim WithEvents ba As BetfairAPI
Dim WithEvents ba2 As BetfairAPI

Sub initBA()
    If ba Is Nothing Then
        Set ba = New BetfairAPI
    End If
End Sub

Sub initBA2()
    If ba2 Is Nothing Then
        Set ba2 = New BetfairAPI
    End If
End Sub

Sub logResult(result As APIResult)
    Select Case result
        Case APIResult.APIResult_OK
            Debug.Print "OK"
        Case APIResult.APIResult_API_ERROR_CHECK_LOG
            Debug.Print "Error. Check log."
        Case APIResult.APIResult_BETFAIR_API_ERROR
            Debug.Print "API Error"
        Case APIResult.APIResult_INVALID_PARAMETERS
            Debug.Print "Invalid parameters"
    End Select
End Sub

Function betTypeEnumName(ByRef betType As BetTypeEnum) As String
    Select Case betType
        Case BetTypeEnum.BetTypeEnum_BACK: betTypeEnumName = "BACK"
        Case BetTypeEnum.BetTypeEnum_LAY: betTypeEnumName = "LAY"
        Case BetTypeEnum.BetTypeEnum_UNKNOWN: betTypeEnumName = "UNKNOWN"
    End Select
End Function

Function betCategoryEnumName(ByRef betCategory As BetCategoryEnum) As String
    Select Case betCategory
        Case BetCategoryEnum.BetCategoryEnum_NORMAL: betCategoryEnumName = "NORMAL"
        Case BetCategoryEnum.BetCategoryEnum_SP: betCategoryEnumName = "SP"
        Case BetCategoryEnum.BetCategoryEnum_SP_WITH_LIMIT: betCategoryEnumName = "SP_WITH_LIMIT"
    End Select
End Function

Function betStatusEnumName(ByRef betStatus As BetStatusEnum) As String
    Select Case betStatus
        Case BetStatusEnum.BetStatusEnum_MATCHED: betStatusEnumName = "MATCHED"
        Case BetStatusEnum.BetStatusEnum_PARTIALLY_MATCHED: betStatusEnumName = "PARTIALLY MATCHED"
        Case BetStatusEnum.BetStatusEnum_UNMATCHED: betStatusEnumName = "UNMATCHED"
    End Select
End Function

Sub testPlaceBets()
    initBA
    ba.init "8000"
    Dim b(1) As Bet, waitForResult As Boolean, ref() As String
    waitForResult = False
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = BetTypeEnum_BACK
    b(0).Price = 500
    b(0).Size = 1
    b(0).token = "token1"
    b(0).keepBetType = KeepBetTypeEnum_KEEP_IN_PLAY
    Set b(1) = New Bet
    b(1).selectionNumber = 1
    b(1).betType = BetTypeEnum_LAY
    b(1).Price = 1.1
    b(1).Size = 1
    b(1).token = "token2"
    b(1).keepBetType = KeepBetTypeEnum_KEEP_IN_PLAY
    If waitForResult Then
        ref = ba.placeBets(b, waitForResult)
        Debug.Print ref(0)
        Debug.Print ref(1)
    Else
        ba.placeBets b, waitForResult
    End If
End Sub

Sub testPlaceBetsMultiBA()
    initBA
    initBA2
    Dim b(0) As Bet, waitForResult As Boolean, ref() As String
    waitForResult = True
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = BetTypeEnum_LAY
    b(0).Price = 1.01
    b(0).Size = 1
    b(0).token = "token1"
    If waitForResult Then
        ref = ba.placeBets(b, waitForResult)
        Debug.Print ref(0)
        ref = ba2.placeBets(b, waitForResult)
        Debug.Print ref(0)
    Else
        ba.placeBets b, waitForResult
        b(0).token = "token2"
        ba2.placeBets b, waitForResult
    End If
End Sub

Private Sub ba_betPlaced(ByVal Bet As ba_com_betfair.IBet)
    Debug.Print "Bet placed fired: Bet ref:" & Bet.ref & "," & betTypeEnumName(Bet.betType) & "," & Bet.resultCode & "," & Bet.token
End Sub

Private Sub ba2_betPlaced(ByVal Bet As ba_com_betfair.IBet)
    Debug.Print Bet.ref & "," & betTypeEnumName(Bet.betType) & "," & Bet.resultCode & "," & Bet.token
End Sub

Sub testUpdateBet()
    initBA
    Dim b(0) As Bet, ref() As String, updateResult As UpdateBetResult
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = BetTypeEnum_BACK
    b(0).Price = 500
    b(0).Size = 1
    ref = ba.placeBets(b, True)
    Set updateResult = ba.updateBet(ref(0), 500, 1, 600, 1)
    Debug.Print "Old ref:" & ref(0) & ", new ref:" & updateResult.newRef & ", size cancelled:" & updateResult.stakeCancelled & ", resultCode:" & updateResult.resultCode
End Sub

Sub testCancelBet()
    initBA
    Dim b(0) As Bet, ref() As String, sizeCancelled As Double
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = BetTypeEnum_BACK
    b(0).Price = 500
    b(0).Size = 1
    ref = ba.placeBets(b, True)
    sizeCancelled = ba.cancelBet(ref(0), ba.marketId)
    Debug.Print "Ref:" & ref(0) & ", size cancelled:" & sizeCancelled
End Sub

Sub testCancelAllBets()
    initBA
    Dim b(0) As Bet, ref() As String, result As APIResult
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = BetTypeEnum_BACK
    b(0).Price = 500
    b(0).Size = 1
    ref = ba.placeBets(b, True)
    Debug.Print ref(0) & " placed."
    result = ba.cancelAllBets
    logResult result
End Sub

Sub testBackAndLayField()
    Dim result As APIResult
    initBA
    result = ba.backField(500, 1)
    logResult result
    result = ba.layField(1.1, 1)
    logResult result
End Sub

Sub testGetBets()
    Dim i As Integer, b As Bet
    initBA
    bets = ba.getBets
    For i = 0 To UBound(bets)
        Set b = bets(i)
        Debug.Print b.ref & "," & betTypeEnumName(b.betType) & "," & b.Price & "," & b.Size & "," & betStatusEnumName(b.betStatus) & "," & b.avgPrice & "," & b.matchedSize & "," & b.remainingSize
    Next
End Sub

Sub testGetBetStatus()
    initBA
    Dim waitForResult As Boolean, bets() As Bet, i As Integer, b As Bet
    waitForResult = False
    bets = ba.getBets
    Dim ref() As String
    ReDim ref(UBound(bets))
    For i = 0 To UBound(bets)
        ref(i) = bets(i).ref
    Next
    If waitForResult Then
        result = ba.getBetStatus(ref, waitForResult)
        For i = 0 To UBound(result)
            Set b = result(i)
            Debug.Print "Ref:" & b.ref
            Debug.Print "Average price matched:" & b.avgPrice
            Debug.Print "Bet type:" & betTypeEnumName(b.betType)
            Debug.Print "Bet status:" & betStatusEnumName(b.betStatus)
        Next
    Else
        ba.getBetStatus ref, waitForResult
    End If
End Sub

Private Sub ba_betsCancelled(bets As Variant, ByVal tabIndex As Long)
    Dim i As Integer, b As Bet
    For i = 0 To UBound(bets)
        Set b = bets(i)
        Debug.Print "Bets cancelled fired: Bet ref:" & b.ref & " cancelled, size cancelled:" & b.sizeCancelled & ",type:" & betTypeEnumName(b.betType)
    Next
End Sub

Private Sub ba_getBetStatusComplete(ByVal bets As Variant)
    For i = 0 To UBound(bets)
        Debug.Print "Get bet status complete fired. Ref:" & bets(i).ref & ",status:" & betStatusEnumName(bets(i).betStatus) & ",type:" & betTypeEnumName(bets(i).betType) & " (async)"
    Next
End Sub

Private Sub ba_betsUpdated(bets As Variant, ByVal betsCount As Long, ByVal tabIndex As Long, ByVal marketId As Long, ByVal marketName As String, ByVal port As String)
    Dim i As Integer, b As Bet
    initBA
    For i = 0 To betsCount - 1
        Set b = bets(i)
        Debug.Print "Bets updated fired: Bet ref:" & b.ref & "," & betTypeEnumName(b.betType) & "," & b.Price & "," & b.Size & "," & betStatusEnumName(b.betStatus) & "," & b.avgPrice & "," & b.matchedSize & "," & b.remainingSize
    Next
End Sub

Sub testProperties()
    initBA
    Debug.Print "BA Version:" & ba.baVersion
    Debug.Print "Commission Rate:" & ba.commissionRate
    Debug.Print "Exchange Rate:" & ba.exchangeRate
    Debug.Print "Include non runners:" & ba.includeNonRunners
    Debug.Print "Interface Type:" & ba.interfaceType
    Debug.Print "Keep Bet Type:" & ba.keepBetType
    Debug.Print "Market Id:" & ba.marketId
    Debug.Print "Market Name:" & ba.marketName
    Debug.Print "Refresh Rate:" & ba.refreshRate
    Debug.Print "Start Time:" & ba.startTime
    Debug.Print "Tabpage Count:" & ba.tabPageCount
    Debug.Print "User Currency:" & ba.userCurrency
End Sub

Sub testAddTabPage()
    initBA
    Debug.Print ba.addTabPage
End Sub

Sub testDeleteTabPage()
    initBA
    Debug.Print ba.deleteTabPage(1)
End Sub

Sub testGetBalance()
    initBA
    Dim bal As Balance
    Set bal = ba.getBalance
    Debug.Print "Balance: " & bal.Balance
    Debug.Print "Available balance:" & bal.availBalance
    Debug.Print "Exposure:" & bal.exposure
End Sub

Sub testGetMetaFunctions()
    initBA
    Dim horse As String
    horse = "Super Superjack"
    Debug.Print "Days since last run:" & ba.getDaysSinceLastRun(horse)
    Debug.Print "Horse form:" & ba.getHorseForm(horse)
    Debug.Print "Saddlecloth:" & ba.getSaddleCloth(horse)
End Sub

Sub testGetSports()
    Dim i As Integer, sports() As BfSport
    initBA
    ba.clearQuickPick 0
    sports = ba.getSports
    For i = 0 To UBound(sports)
        If sports(i).sport = "Horse Racing" Then
            testGetEvents sports(i).sportId, sports(i).sport
        End If
    Next
End Sub

Sub testGetEvents(id As Long, path As String)
    Dim i As Integer, evnts() As BfEvent
    evnts = ba.getEvents(id)
    For i = 0 To UBound(evnts)
        If evnts(i).isMarket Then
            If InStr(path, "/GB") <> 0 Then
                ba.addMarketToQuickPick evnts(i).eventId, 0
                Debug.Print evnts(i).eventName
            End If
        Else
            testGetEvents evnts(i).eventId, path & "/" & evnts(i).eventName
        End If
    Next
End Sub

Sub testGetAllTradedVolume()
    Dim i As Integer, j As Integer, tvs() As TradedVolumeSelection, tv() As TradedVolume
    Dim waitForResult As Boolean
    waitForResult = False
    initBA
    If waitForResult Then
        tvs = ba.getAllTradedVolume(waitForResult)
        For i = 0 To UBound(tvs)
            tv = tvs(i).tradedVolumes
            For j = 0 To UBound(tv)
                Debug.Print tvs(i).selectionName & "," & tv(j).Price & ","; tv(j).totalMatchedAmount
            Next
        Next
    Else
        ba.getAllTradedVolume waitForResult
    End If
End Sub

Private Sub ba_getAllTradedVolumeComplete(ByVal TradedVolume As Variant, ByVal resultCode As String)
    Dim tv As TradedVolume
    For i = 0 To UBound(TradedVolume)
        tvs = TradedVolume(i).tradedVolumes
        For j = 0 To UBound(tvs)
            Set tv = tvs(j)
            Debug.Print resultCode & "," & TradedVolume(i).selectionName & "," & tv.Price & ","; tv.totalMatchedAmount
        Next
    Next
End Sub

Sub testGetAllMarketDepth()
    Dim i As Integer, j As Integer, mds() As MarketDepthSelection, md() As MarketDepth
    Dim waitForResult As Boolean
    waitForResult = True
    initBA
    If waitForResult Then
        mds = ba.getMarketDepth(waitForResult)
        For i = 0 To UBound(mds)
            md = mds(i).prices
            For j = 0 To UBound(md)
                Debug.Print mds(i).selectionName & "," & md(j).Price & ","; md(j).backAmountAvailable
            Next
        Next
    Else
        ba.getMarketDepth waitForResult
    End If
End Sub

Private Sub ba_getMarketDepthComplete(ByVal MarketDepth As Variant, ByVal resultCode As String)
    Dim md As MarketDepth
    For i = 0 To UBound(MarketDepth)
        prices = MarketDepth(i).prices
        For j = 0 To UBound(prices)
            Set md = prices(j)
            Debug.Print resultCode & "," & MarketDepth(i).selectionName & "," & md.Price & ","; md.backAmountAvailable
        Next
    Next
End Sub

Private Sub testGetMarketDepthString()
    Dim md As String
    initBA
    md = ba.getMarketDepthString
    Debug.Print md
End Sub

Private Sub testGetMetaData()
    initBA
    Dim md As SelectionMetaData
    Set md = ba.getMetaData("Super Superjack")
    Debug.Print "Age/Weight:" & md.ageWeight
    Debug.Print "Bred:" & md.bred
    Debug.Print "Colour/Sex:" & md.colourSex
    Debug.Print "Dam:" & md.dam
    Debug.Print "Dam/Sire:" & md.damSire
    Debug.Print "Days since last run:" & md.daysSinceLastRun
    Debug.Print "Forecast price:" & md.forecastPrice
    Debug.Print "Form:" & md.form
    Debug.Print "Jockey:" & md.jockey
    Debug.Print "Jockey claim:" & md.jockeyClaim
    Debug.Print "Official rating:" & md.officialRating
    Debug.Print "Owner:" & md.owner
    Debug.Print "Saddlecloth:" & md.saddleCloth
    Debug.Print "Selection:" & md.Selection
    Debug.Print "Sire:" & md.sire
    Debug.Print "Stall draw:" & md.stallDraw
    Debug.Print "Trainer:" & md.trainer
    Debug.Print "Wearing:" & md.wearing
End Sub

Function RightPad(str As String, totalLength As Integer) As String
    Dim paddingLength As Integer
    paddingLength = totalLength - Len(str)
    If paddingLength > 0 Then
        RightPad = str & Space(paddingLength)
    Else
        RightPad = str
    End If
End Function

Private Sub testGetPrices()
    Dim prices() As Price, i As Integer
    initBA
    prices = ba.getPrices
    For i = 0 To UBound(prices)
        Debug.Print RightPad(prices(i).Selection, 50) & vbTab & RightPad(prices(i).backOdds1, 12) & vbTab & RightPad(prices(i).layOdds1, 12)
    Next
End Sub

Private Sub ba_marketSettled(bets As Variant)
    Dim i As Integer, b As Bet
    For i = 0 To UBound(bets)
        Set b = bets(i)
        Debug.Print "Market settled fired: Selection name: " & b.selectionName & "," & betTypeEnumName(b.betType) & "," & b.avgPrice & "," & b.Size & "," & b.result
    Next
End Sub

'Private Sub ba_pricesUpdated(prices As Variant, ByVal pricesCount As Long, ByVal tabIndex As Long, ByVal marketId As Long, ByVal marketName As String, ByVal port As String)
    'Dim i As Integer, p As Price
    'For i = 0 To pricesCount - 1
    '    Set p = prices(i)
    '    Debug.Print RightPad(p.Selection, 50) & vbTab & RightPad(p.backOdds1, 12) & vbTab & RightPad(p.layOdds1, 12)
    'Next
'End Sub

Private Sub testGetUserName()
    initBA
    Debug.Print ba.getUserName
End Sub

Sub testClearQuickPick()
    Dim result As APIResult
    initBA
    result = ba.clearQuickPick(0)
    logResult result
End Sub

Private Sub testLoadQuickPick()
    Dim result As APIResult
    initBA
    result = ba.loadQuickPickList(QuickPickMarketType_UK_HORSE_RACING_WIN, True)
    logResult result
End Sub

Private Sub testOpenFirstQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openFirstQuickPickMarket
    logResult result
End Sub

Private Sub testOpenNextQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openNextQuickPickMarket
    logResult result
End Sub

Private Sub testOpenPreviousQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openPreviousQuickPickMarket
    logResult result
End Sub

Private Sub testOpenLastQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openLastQuickPickMarket
    logResult result
End Sub

Private Sub testOpenMarket()
    Dim result As APIResult, marketId As Long
    initBA
    marketId = ba.marketId
    testOpenLastQuickPickMarket
    result = ba.openMarket(marketId)
    logResult result
End Sub

Private Sub testRefreshMarkets()
    Dim result As APIResult
    initBA
    result = ba.refreshMarkets
    logResult result
End Sub

Private Sub testSetQuickPickListAutoSelect()
    Dim result As APIResult
    initBA
    result = ba.setQuickPickAutoSelect(False, -1)
    logResult result
End Sub

Private Sub testProcessExcelTriggers()
    initBA
    ba.processExcelTriggers ThisWorkbook.Name, Me.Name
End Sub

Private Sub testGetTPD()
    Dim tpd() As TpdSelection, i As Integer
    initBA
    tpd = ba.getTpd
    For i = 0 To UBound(tpd)
        Set runner = tpd(i)
        Debug.Print "Timestamp:" & runner.timestamp
        Debug.Print "Saddle cloth:" & runner.saddleCloth
        Debug.Print "Market id:" & runner.marketId
        Debug.Print "Selection id:" & runner.selectionId
        Debug.Print "Active:" & runner.Active
        Debug.Print "Cadence error:" & runner.cadenceError
        Debug.Print "Distance to go in metres:" & runner.distanceToGoInMetres
        Debug.Print "Fastest speed:" & runner.fastestSpeed
        Debug.Print "Leader direction:" & runner.leaderDirection
        Debug.Print "Leader distance:" & runner.leaderDistance
        Debug.Print "Leader speed:" & runner.leaderSpeed
        Debug.Print "Metres back from leader:" & runner.metresBackFromLeader
        Debug.Print "Progress:" & runner.progress
        Debug.Print "Running order:" & runner.runningOrder
        Debug.Print "Running time in seconds:" & runner.runningTimeInSeconds
        Debug.Print "Slowest speed:" & runner.slowestSpeed
        Debug.Print "Speed:" & runner.speed
        Debug.Print "Speed (course par):" & runner.speedCoursePar
        Debug.Print "Speed ranking:" & runner.speedRanking
        Debug.Print "Stride frequency:" & runner.strideFrequency
        Debug.Print "Stride frequency (course par):" & runner.strideFrequencyCoursePar
        Debug.Print "Total distance in metres:" & runner.totalDistanceInMetres
        Debug.Print "Velocity error:" & runner.velocityError
        Debug.Print "Velocity fluctuation:" & runner.VelocityFluctuation
    Next
    Debug.Print
End Sub
