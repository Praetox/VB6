VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   11668
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const timeout = 7

Private Sub Command1_Click()
On Error Resume Next
user = InputBox("Start spam at user #", "NordicSpammer", "7800")
wb.Navigate ("http://www.nordicmafia.net/nordic/index.php?side=online&start=" & user)
Call slp
names = Split(Split(wb.Document.All.Item(0).innertext, " spillere" & vbNewLine & vbNewLine & vbNewLine)(1), ", ")
Open "c:\nmspam_tmp.dat" For Output As #1
For a = 0 To 599
    Print #1, names(a)
Next a
Close #1
DoEvents
MsgBox "Read names."
Open "c:\nmspam_tmp.dat" For Input As #1
done = 0
While Not EOF(1)
    Input #1, tmp
    If Not EOF(1) Then
        wb.Navigate ("about:" & _
            "<form name=""form1"" method=""post"" action=""http://www.nordicmafia.net/nordic/index.php?side=pm_ny2"">" & _
            "<input type=""text"" maxlength=""55"" name=""til"" value="""">" & _
            "<input type=""text"" maxlength=""50"" name=""tittel"" value="""">" & _
            "<textarea name=""melding"" wrap=""VIRTUAL"" cols=""50"" rows=""7""></textarea>" & _
            "<input type=""submit"" name=""Submit5222"" value=""Send"">" & _
            "</form>")
        DoEvents
        wb.Document.All("til").Value = tmp
        wb.Document.All("tittel").Value = "Lei av manuell ranking?"
        wb.Document.All("melding").Value = "" & _
            "Er du lei av monoton klikking dag ut og dag inn? Kunne du gitt hva som helst for en som kan spille " & _
            "for deg? Da er dette lykkedagen din. Jeg har avsatt en del tid til og lage et dataprogram som " & _
            "ranker for deg. Du trenger bare starte programmet, og etter et par klikk vil det spille automatisk." & _
            vbCrLf & vbCrLf & _
            "Tror du meg ikke? Sjekk denne youtube filmen. http://www.youtube.com/watch?v=iqj0IYCtqYc" & _
            vbCrLf & vbCrLf & _
            "Hvis du vil laste den ned er det bare og navigere til http://nordic.awardspace.com" & _
            vbCrLf & vbCrLf & _
            "Dine ensformige tider er over. Nyt nordics fremtid! =D"
        DoEvents
        wb.Document.All("Submit5222").Click
        Call slp
        For a = 0 To 19
            Sleep (1000)
            Form1.Caption = 19 - a & " - " & done
            DoEvents
        Next a
        done = done + 1
    End If
Wend
Close #1
MsgBox "Spammed all! =D"
End Sub

Private Sub Form_Load()
'Scan for start: Split(wb.Document.All.Item(0).innertext, " spillere" & vbNewLine & vbNewLine & vbNewLine)(1)
'Scan for playername: mid$(wb.Document.All.Item(0).innertext, InStr(1, wb.Document.All.Item(0).innertext, "- G -") - 50, 100)
End Sub

Private Sub slp()
    DoEvents
    t = Timer
    While wb.Busy And (Timer - t) < timeout
        Sleep (1)
        DoEvents
    Wend
    DoEvents
End Sub
