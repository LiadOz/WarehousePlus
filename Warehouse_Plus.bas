Attribute VB_Name = "Warehouse_Plus"
' References needed: Microsoft Internet Controls, Microsoft HTML Object Library
' Also needed the Utility module for CutLeftUntil and CutRightUntil methods
' A module scrape information from a stock managing site
' Specific information removed
' This module is comprised of 4 parts:
' 1. Main Site Navigation, are methods which help going into different menus including logging and opening the site.
' 2. Navigation Buttons, used to travel inside the menus.
' 3. Utility Functions, used to handle the non scraping needs
' 4. Main Scraping Functions, the functions which do the scraping and parsing the final data

Declare Sub Sleep Lib "kernel132" (ByVal dwMilliseconds As Long)
Dim IE As InternetExplorer
Dim currWin As HTMLDocument

Sub console()
' Console to test scraping paths

getIE

' do tests below


End Sub

' -------------------------------------------------------------
' -------------------------------------------------------------
' -----------------  Main Scraping Functions  -----------------
' -------------------------------------------------------------
' -------------------------------------------------------------

Sub initStockCollection()
' Goes to the screen from which scraping can begin

openSite
login
nituv ("menu")

End Sub


Function getStockCollection(ct As String) As Collection
' This method uses getAllStock and stockInAllWarehouses to create a readable collection of part stock


defaultMenu ct, "type a"

Dim col As Collection
Dim allStock As sting

allStock = getAllStock ' Stock in all warehouses
Set col = stockInWarehouses ' Stock in specific warehouses


If col.Count > 0 Then
' Putting the all stock first if there is even in warehouses
    col.Add "омай-:" & allStock, , 1
End If

Set getStockCollection = col

End Function

Function getAllStock() As String
' Returns stock in all warehouses as string, can work only in specific screens

Set currWin = IE.Document.frames("the_frame_name").Document
Dim table As IHTMLElement
Set table = currWin.getElementsByName("table-tag")("index of specific table")

wait

getAllStock = "0"
On Error Resume Next
getAllStock = CutLeftUntil(table.Children("specific child").innerText, ".", True)

End Function


Function stockInWarehouses() As Collection
' Returns a collection of stocks in specific warehouses, can only work in specific screens
' The information is stored in the specific format: "warehouse_id-count"
' For example warehouse_42-420 (warehouse_42 is the warehouse, 420 is the item count)


startAgain:
pushForward

Dim stock As New Collection
Dim table As IHTMLElement
Dim warehouse As String
Dim amount As String

Set currWin = IE.Document.frames("the_frame_name").Document

On Error GoTo subEnd
Set table = currWin.all("specific route to the correct table")

For tRow = 2 To (table.Children.Length - 1)
    Set rowElement = table.Children(tRow).Children(0)
    warehouse = table.Children(tRow).Children("route")
    amount = table.Children(tRow).Children("another route")
    
    If warehouse = "" Then
        ' If the end of current table reached
        GoTo nextIteration
    End If
    
    If Left(warehouse, 1) = "*" Then
        ' Formatting the entry
        stock.Add (CutRightUntil(warehouse, "*", True, 3) & ":-" & CutLeftUntil(amount, ".", True))
    End If
nextIteration:
Next
GoTo startAgain

subEnd:

pushEnter
Set stockInWarehouse = stock

End Function

Sub getStockInExcel()

' Gets a list of parts from selection and retrievs the stock in the site for them in the cell to the left

Dim myRange As range
Set myRange = Application.Selection
Dim aCell As range
Dim ct As String
Dim mystr As sting
Dim stockcol As Collection

Dim errcol As New Collection

initStockCollection

For Each aCell In myRange

    ct = aCell
    errcol.Add ct
    Sleep 100
    Set stockcol = getStockCollection(ct)
    ' Converting the format to a suitable one
    For Each st In stockcol
        mystr = st
        Set aCell = aCell.Offset(, 1)
        aCell = "'" & CutLeftUntil(mystr, "-", True) ' Offsetting the cursor to the left of the cell
        Set aCell = aCell.Offset(, 1)
        aCell = CInt(CutRightUntil(mystr, "-", True))
        
    Next st
    
Next aCell

Dim ans As Variant

' Error debugging
ans = "no"
ans = InputBox("Debug errors", "Code Finished")
If ans = yes Then
    For Each st In errcol
        MsgBox st
    Next st
End If

endAll
        
End Sub

' ------------------------------------------------------------
' ------------------------------------------------------------
' ------------------- Main Site Navigation -------------------
' ------------------------------------------------------------
' ------------------------------------------------------------

Sub openSite()
' Opens the site as a new instance of IE the window will not be visible

Set IE = CreateObject("InternetExplorer.Application")
IE.Silent = True
IE.Navigate "http://www.Warehouse_site/portal/default.asap"
IE.Visible = False

wait

End Sub
Sub openVisSite()
' Opens the site as a new instance of IE the window will be visible

Set IE = CreateObject("InternetExplorer.Application")
IE.Silent = True
IE.Navigate "http://www.Warehouse_site/portal/default.asap"
IE.Visible = True

wait

End Sub

Sub login()
' Logs in to the site using username and password

Set currWin = IE.Document
currWin.getElementById("user-id").Value = "username"
currWin.getElementById("password-id").Value = "password"
currWin.getElementById("button-id").Click

wait

End Sub


Sub nituv(menu As String)
' Goes to a certain menu according to the menu option

Set currWin = IE.Document
Set currWin = currWin.frames("the_frame_name").Document
currWin.getElementById("nituv-id")(0).Value = menu
currWin.getElementsByTagName("next-tag")(2).Click

wait

End Sub

Sub enterCT(ct As String)
' Enter search for ct

Set currWin = IE.Document
Set currWin = currWin.frames("the_frame_name").Document

currWin.getElementsByName("ct")(0).Value = ct

End Sub


Sub enterPN(pn As String)
' Enter search for pn

Set currWin = IE.Document
Set currWin = currWin.frames("the_frame_name").Document

currWin.getElementsByName("pn")(0).Value = pn

End Sub


Sub selectPartType(ptype As String)
' Select a category of the part to search

Dim optionNum As Interger

Select Case ptype
    Case "type a"
        optionNum = 1
    Case "type b"
        optionNum = 2
    Case "type c"
        optionNum = 3
    Case Else
        optionNum = 1
End Select

Set currWin = IE.Document
Set currWin = currWin.frames("the_frame_name").Document

currWin.getElementsByName("option")(0).Value = optionNum

End Sub


Sub defaultMenu(ct As String, ptype As String)
' Entering the most frequent menu

enterCT ct
selectPartType ptype

wait
pushEnter
wait

End Sub

' -------------------------------------------------------------
' -------------------------------------------------------------
' -------------------   Navigation Buttons  -------------------
' -------------------------------------------------------------
' -------------------------------------------------------------

Sub pushPrevious()
' Goes to previous in the site's interface if exists

Set currWin = IE.Document
Set currWin = currWin.frames("buttons-frame").Document

currWin.all("prev_button-index").Click
wait

End Sub

Sub pushBack()
' Goes back in the site's interface if exists

Set currWin = IE.Document
Set currWin = currWin.frames("buttons-frame").Document

currWin.all("back_button-index").Click
wait

End Sub

Sub pushForward()
' Goes forward in the site's interface if exists

Set currWin = IE.Document
Set currWin = currWin.frames("buttons-frame").Document

currWin.all("forward_button-index").Click
wait

End Sub

Sub pushNext()
' Goes to next in the site's interface if exists

Set currWin = IE.Document
Set currWin = currWin.frames("buttons-frame").Document

currWin.all("next_button-index").Click
wait

End Sub

Sub pushEnter()
' Goes to previous in the site's interface if exists

Set currWin = IE.Document
Set currWin = currWin.frames("buttons-frame").Document

currWin.all("enter_button-index").Click
wait

End Sub

' ------------------------------------------------------------
' ------------------------------------------------------------
' --------------------  Utility Functions --------------------
' ------------------------------------------------------------
' ------------------------------------------------------------

Sub endAll()

IE.Quit
Set currWin = Nothing

End Sub

Sub wait()
' Waiting for site to load needs to be called after each menu change or button press

Sleep 100
Do While IE.Busy Or IE.ReadyState <> 4: Sleep 100: Loop
Sleep 100
Do While IE.Busy Or IE.ReadyState <> 4: Sleep 100: Loop

End Sub

Sub getIE()
' Defines the IE object as the current open window of the site by looking at each open IE instance
' If none exist a new IE window will be initiated

Dim shellWins As ShellWindows
Set shellWins = New ShellWindows

Dim processName As String

For i = 0 To shellWins.Count - 1
    processName = ""
    On Error Resume Next
    processName = shellWins.Item(i).Name
    If processName = "Internet Explorer" Then
        If shellWins.Item(i).LocationURL = "http://www.Warehouse_site/portal/default.asap" Then
            Set IE = shellWins.Item(i)
            Exit Sub
        End If
    End If
Next

' If there are no active windows of the site

Set IE = CreateObject("InternetExplorer.Application")
IE.Silent = True
IE.Navigate "http://www.Warehouse_site/portal/default.asap"
IE.Visible = True

Set shellWins = Nothing

wait

End Sub