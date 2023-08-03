Dim WithEvents ba As BetfairAPI

Sub initBA()
    If ba Is Nothing Then
        Set ba = New BetfairAPI
    End If
End Sub

Sub testPlaceBet()
    initBA
    Dim b As New Bet, waitForResult As Boolean, ref As String
    waitForResult = False
    b.selectionNumber = 0
    b.betType = "B"
    b.Price = 500
    b.Size = 1
    b.token = "token1"
    If waitForResult Then
        ref = ba.placeBet(b, waitForResult)
        Debug.Print ref
    Else
        ba.placeBet b, waitForResult
    End If
End Sub

Sub testPlaceBets()
    initBA
    Dim b(1) As Bet, waitForResult As Boolean, ref() As String
    waitForResult = False
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = "B"
    b(0).Price = 500
    b(0).Size = 1
    b(0).token = "token1"
    Set b(1) = New Bet
    b(1).selectionNumber = 1
    b(1).betType = "L"
    b(1).Price = 1.1
    b(1).Size = 1
    b(1).token = "token2"
    If waitForResult Then
        ref = ba.placeBets(b, waitForResult)
        Debug.Print ref(0)
        Debug.Print ref(1)
    Else
        ba.placeBets b, waitForResult
    End If
End Sub

Private Sub ba_betPlaced(ByVal Bet As BA_COM_Betfair.IBet)
    Debug.Print Bet.ref & "," & Bet.resultCode & "," & Bet.token
End Sub

Sub testUpdateBet()
    initBA
    Dim b As New Bet, ref As String, updateResult As UpdateBetResult
    b.selectionNumber = 0
    b.betType = "B"
    b.Price = 500
    b.Size = 1
    ref = ba.placeBet(b, True)
    Set updateResult = ba.updateBet(ref, 500, 1, 600, 1)
    Debug.Print "Old ref:" & ref & ", new ref:" & updateResult.newRef & ", resultCode:" & updateResult.resultCode
End Sub

Sub testCancelBet()
    initBA
    Dim b As New Bet, ref As String, sizeCancelled As Double
    b.selectionNumber = 0
    b.betType = "B"
    b.Price = 500
    b.Size = 1
    ref = ba.placeBet(b, True)
    sizeCancelled = ba.cancelBet(ref, ba.marketId)
    Debug.Print "Ref:" & ref & ", size cancelled:" & sizeCancelled
End Sub

' trans log show cancel all bets clicked. should show it was a com operation.
Sub testCancelAllBet()
    initBA
    Dim b As New Bet, ref As String, result As String
    b.selectionNumber = 0
    b.betType = "B"
    b.Price = 500
    b.Size = 1
    ref = ba.placeBet(b, True)
    result = ba.cancelAllBets
    Debug.Print result
End Sub

Sub testGetBet()
    initBA
    Dim b As New Bet, b1 As Bet, ref As String, result As String
    waitForResult = False
    b.selectionNumber = 0
    b.betType = "B"
    b.Price = 500
    b.Size = 1
    ref = ba.placeBet(b, True)
    Set b1 = ba.getBet(ref)
    Debug.Print b1.ref & ","; b1.matched
End Sub

Sub testGetBets()
    initBA
    Dim b(1) As Bet, i As Integer, bets() As Bet
    waitForResult = False
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = "B"
    b(0).Price = 500
    b(0).Size = 1
    b(0).token = "token1"
    Set b(1) = New Bet
    b(1).selectionNumber = 1
    b(1).betType = "L"
    b(1).Price = 1.1
    b(1).Size = 1
    b(1).token = "token2"
    ba.placeBets b, True
    bets = ba.getBets
    For i = 0 To 1
        Debug.Print bets(i).ref & ","; bets(i).betType & ","; bets(i).Price & "," & bets(i).Size
    Next
End Sub

Sub testBackAndLayField()
    initBA
    Debug.Print ba.backField(500, 1)
    Debug.Print ba.layField(1.1, 1)
End Sub

Sub testGetBetStatus()
    initBA
    Dim inBackground As Boolean
    inBackground = True
    Dim b(1) As Bet, ref() As String, result() As Bet
    Set b(0) = New Bet
    b(0).selectionNumber = 0
    b(0).betType = "B"
    b(0).Price = 500
    b(0).Size = 1
    Set b(1) = New Bet
    b(1).selectionNumber = 1
    b(1).betType = "L"
    b(1).Price = 1.1
    b(1).Size = 1
    ref = ba.placeBets(b, True)
    If Not inBackground Then
        result = ba.getBetStatus(ref, inBackground)
        For i = 0 To 1
            Debug.Print result(i).ref & "," & result(i).betStatus
        Next
    Else
        ba.getBetStatus ref, inBackground
    End If
End Sub

Private Sub ba_betsCancelled(bets As Variant, ByVal tabIndex As Long)
    Dim i As Integer, b As Bet
    For i = 0 To UBound(bets)
        Set b = bets(i)
        Debug.Print "Bet ref:" & b.ref & " cancelled, size cancelled:" & b.sizeCancelled
    Next
End Sub

Private Sub ba_getBetStatusComplete(ByVal bets As Variant)
    For i = 0 To UBound(bets)
        Debug.Print bets(i).ref & "," & bets(i).betStatus & " (async)"
    Next
End Sub

Private Sub ba_betsUpdated(bets As Variant, ByVal betsCount As Long, ByVal tabIndex As Long, ByVal marketId As Long, ByVal marketName As String, ByVal port As String)
    Dim i As Integer, b As Bet
    For i = 0 To betsCount - 1
        Set b = bets(i)
        Debug.Print RightPad(b.selectionName, 50) & vbTab & RightPad(b.betType, 12) & vbTab & RightPad(b.matched, 12)
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

Sub testClearQuickPick()
    initBA
    Debug.Print ba.clearQuickPick(0)
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
    Debug.Print "Days since last run:" & ba.getDaysSinceLastRun("Songo")
    Debug.Print "Horse form:" & ba.getHorseForm("Songo")
    Debug.Print "Saddlecloth:" & ba.getSaddleCloth("Songo")
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
    waitForResult = False
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
    Set md = ba.getMetaData("Westerton")
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

Private Sub ba_pricesUpdated(prices As Variant, ByVal pricesCount As Long, ByVal tabIndex As Long, ByVal marketId As Long, ByVal marketName As String, ByVal port As String)
    'Dim i As Integer, p As Price
    'For i = 0 To pricesCount - 1
    '    Set p = prices(i)
    '    Debug.Print RightPad(p.Selection, 50) & vbTab & RightPad(p.backOdds1, 12) & vbTab & RightPad(p.layOdds1, 12)
    'Next
End Sub

Private Sub testGetUserName()
    initBA
    Debug.Print ba.getUserName
End Sub

Private Sub testLoadQuickPick()
    Dim result As APIResult
    initBA
    result = ba.loadQuickPickList(QuickPickMarketType_UK_HORSE_RACING_WIN, True)
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testOpenFirstQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openFirstQuickPickMarket
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testOpenNextQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openNextQuickPickMarket
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testOpenPreviousQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openPreviousQuickPickMarket
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testOpenLastQuickPickMarket()
    Dim result As APIResult
    initBA
    result = ba.openLastQuickPickMarket
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testRefreshMarkets()
    Dim result As APIResult
    initBA
    result = ba.refreshMarkets
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

Private Sub testSetQuickPickListAutoSelect()
    Dim result As APIResult
    initBA
    result = ba.setQuickPickAutoSelect(False, -1)
    If result = APIResult_OK Then
        Debug.Print "OK"
    Else
        Debug.Print "ERROR"
    End If
End Sub

