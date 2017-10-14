Attribute VB_Name = "modLang"
Public Lng As Language, lNor As Language, lEng As Language
Type Language
    slbInbox As String
    slbOutbox As String
    slbAppStatus As String
    sinStats As String
    soutStats As String
    sfrmAccessInf As String
    sfrmGenConf As String
    sfrmContList As String
    slbUsername As String
    slbPassword As String
    saccUser_ As String
    saccPass_ As String
    saccSave As String
    saccSave_ As String
    slbMaxread As String
    slbMaxread_ As String
    slbAutosync As String
    slbAutosync_ As String
    slbAutocheck As String
    slbAutocheck_ As String
    slbCntRefresh As String
    slbCntRefresh_ As String
    slbCntWarn As String
    slbCntWarn_ As String
    slbCntLaunch As String
    slbCntLaunch_ As String
    scfgSave As String
    scfgSave_ As String
    sinFilter0 As String
    sinFilter0_ As String
    sinFilter1 As String
    sinFilter1_ As String
    sinFilter2 As String
    sinFilter2_ As String
    sinDelThese As String
    sinDelThese_ As String
    sinDelAll As String
    sinDelAll_ As String
    scmdRefresh As String
    scmdRefresh_ As String
    scmdFilter As String
    scmdFilter_ As String
    scmdDeleteRead As String
    scmdDeleteRead_ As String
    scmdDelBy As String
    scmdDelBy_ As String
    scmdNew As String
    scmdNew_ As String
    scmdAdvertise As String
    scmdAdvertise_ As String
    scntAdd As String
    scntAdd_ As String
    scntRem As String
    scntRem_ As String
    scntSendPM As String
    scntSendPM_ As String
    scntRefresh As String
    scntRefresh_ As String
    slbSendCurr As String
    slbSendCurr_ As String
    slbMsgComp As String
    sCompSend0 As String
    sCompSend0_ As String
    sCompSend1 As String
    sCompSend1_ As String
    slbMsgView As String
    sViewUnread0 As String
    sViewUnread1 As String
    sViewReply As String
    sViewReply_ As String
    sViewSave As String
    sViewSave_ As String
    sViewDel As String
    sViewDel_ As String
    
    val As String
    sMsgDelThese As String
    sMsgDelAll As String
    sMsgDelRead As String
    sMsgAddCnt As String
    sMsgRemCnt As String
    sMsgAdvSure As String
    sMsgDelThisMsg As String
End Type

Sub Lang_Init()
    lEng.val = "ENG"
    lEng.saccPass_ = "Your nordicmafia password."
    lEng.saccUser_ = "Your nordicmafia username."
    lEng.saccSave = "Save this information"
    lEng.saccSave_ = "Make Postal remember your username and password. Required for ""Check for PMs at launch"" to work."
    lEng.scfgSave = "Save configuration"
    lEng.scfgSave_ = "Make Postal remember this configuration."
    lEng.scmdAdvertise = "Become a Postal advertiser!"
    lEng.scmdAdvertise_ = "Advertise for Postal by sending PMs to 1000 online players."
    lEng.scmdDeleteRead = "Delete read"
    lEng.scmdDeleteRead_ = "Delete all the PMs that you have read."
    lEng.scmdDelBy = "Delete by..."
    lEng.scmdDelBy_ = "Delete messages by topic or author name, for example all fight club messages."
    lEng.scmdFilter = "Filter"
    lEng.scmdFilter_ = "Filter your messages. Same as the ""Show ..."" buttons in the inbox field."
    lEng.scmdNew = "New"
    lEng.scmdNew_ = "Show the message composer, prepared for a new message."
    lEng.scmdRefresh = "Refresh"
    lEng.scmdRefresh_ = "Refresh the list over messages, and download any new messages to Postal."
    lEng.scntAdd = "Add"
    lEng.scntAdd_ = "Add a new person to the contacts list."
    lEng.scntRem = "Remove"
    lEng.scntRem_ = "Remove the selected person from the contacts list."
    lEng.scntSendPM = "Send PM"
    lEng.scntSendPM_ = "Send a PM to the selected person on the contacts list."
    lEng.scntRefresh = "Refresh list"
    lEng.scntRefresh_ = "Refresh the contacts list."
    lEng.sCompSend0 = "Add to queue"
    lEng.sCompSend0_ = "Send the message with priority level two."
    lEng.sCompSend1 = "Send ASAP"
    lEng.sCompSend1_ = "Send the message with priority level one."
    lEng.sfrmAccessInf = "Nordicmafia Access Information"
    lEng.sfrmGenConf = "General configuration"
    lEng.sfrmContList = "Contacts list"
    lEng.sinDelAll = "Delete ALL messages"
    lEng.sinDelAll_ = "Delete all messages from your inbox."
    lEng.sinDelThese = "Delete these messages"
    lEng.sinDelThese_ = "Delete the 15 newest messages in your inbox (the first page on nordicmafia)."
    lEng.sinFilter0 = "Show all"
    lEng.sinFilter0_ = "Show all messages."
    lEng.sinFilter1 = "Show unread"
    lEng.sinFilter1_ = "Show only the messages that has yet to be read."
    lEng.sinFilter2 = "Show read"
    lEng.sinFilter2_ = "Show only the messages that you have already opened."
    lEng.sinStats = "%1 total, %2 unread"
    lEng.slbAppStatus = "Application status"
    lEng.slbAutocheck = "Check for PMs at launch"
    lEng.slbAutocheck_ = "Automatically fetch all PMs when you start Postal. Recommended."
    lEng.slbAutosync = "Automatic PM checking"
    lEng.slbAutosync_ = "Check for new PMs every x second. Put 0 to disable the feature."
    lEng.slbCntLaunch = "Get contacts at launch"
    lEng.slbCntLaunch_ = "Automatically fetch your contacts' online statuses when you launch Postal."
    lEng.slbCntRefresh = "Online list refresher"
    lEng.slbCntRefresh_ = "Update the contact list every x second. Put 0 to disable the feature."
    lEng.slbCntWarn = "Contact state warner"
    lEng.slbCntWarn_ = "Warns you if a contact logs on or off."
    lEng.slbInbox = ".:: Inbox ::."
    lEng.slbMaxread = "Max messages to read"
    lEng.slbMaxread_ = "The max amount of messages that Postal should read. Put 0 to read all 15."
    lEng.slbMsgComp = "Message composer"
    lEng.slbMsgView = "Message viewer"
    lEng.slbOutbox = ".:: Outbox ::."
    lEng.slbPassword = "Password"
    lEng.slbSendCurr = "Send curr. PM"
    lEng.slbSendCurr_ = "Click a contact, and the last written PM will be sent to him/her also."
    lEng.slbUsername = "Username"
    lEng.soutStats = "%1 queued. Next: %2"
    lEng.sViewDel = "Delete"
    lEng.sViewDel_ = "Delete this message from both nordicmafia and Postal. If you saved this message using the ""Save"" button, the saved version will still remain."
    lEng.sViewReply = "Reply"
    lEng.sViewReply_ = "Click here to open the message composer ready for replying."
    lEng.sViewSave = "Save"
    lEng.sViewSave_ = "Save this message to the application's folder."
    lEng.sViewUnread0 = "New message"
    lEng.sViewUnread1 = "Old message"
    
    lEng.sMsgDelThese = "Are you sure that you want to delete the messages currently shown in inbox?"
    lEng.sMsgDelRead = "Are you sure that you want to delete all read messages?"
    lEng.sMsgDelAll = "Are you sure that you want to delete all your messages?"
    lEng.sMsgAddCnt = "Please enter the nordicmafia name of the person you want to add."
    lEng.sMsgRemCnt = "Are you sure that you want to delete ""%1"" from your contacts list?"
    lEng.sMsgAdvSure = "Are you sure you want to do that? Your account will most likely be banished."
    lEng.sMsgDelThisMsg = "Are you sure you want to delete the following message?"
    
    lNor.val = "NOR"
    lNor.saccPass_ = "Ditt nordicmafia passord."
    lNor.saccUser_ = "Ditt nordicmafia brukernavn."
    lNor.saccSave = "Lagre brukernavn/passord"
    lNor.saccSave_ = "Gjør at Postal husker brukernavn og passord. Kreves for at ""Hent PMs v/ oppstart"" skal fungere."
    lNor.scfgSave = "Lagre konfigurasjon"
    lNor.scfgSave_ = "Gjør at Postal husker oppsettet."
    lNor.scmdAdvertise = "Reklamer for Postal!"
    lNor.scmdAdvertise_ = "Reklamerer for Postal ved å sende PMs til 1000 påloggede spillere."
    lNor.scmdDeleteRead = "Slett leste"
    lNor.scmdDeleteRead_ = "Slett alle meldinger som du har lest."
    lNor.scmdDelBy = "Slett etter..."
    lNor.scmdDelBy_ = "Slett meldinger med spesifikke emner eller avsendere, f.eks. alle fight club meldinger."
    lNor.scmdFilter = "Filter"
    lNor.scmdFilter_ = "Filtrér meldinger; samme som ""Vis ..."" knappene i innboks-feltet."
    lNor.scmdNew = "Ny"
    lNor.scmdNew_ = "Skriv en ny PM."
    lNor.scmdRefresh = "Oppdater"
    lNor.scmdRefresh_ = "Sjekker etter nye meldinger, og laster dem ned for visning."
    lNor.scntAdd = "Legg til"
    lNor.scntAdd_ = "Legger til en person til kontaktlisten."
    lNor.scntRem = "Fjern"
    lNor.scntRem_ = "Fjerner en person fra kontaktlisten."
    lNor.scntSendPM = "Send PM"
    lNor.scntSendPM_ = "Sender en PM til personen du valgte i kontaktlisten."
    lNor.scntRefresh = "Oppdater"
    lNor.scntRefresh_ = "Oppdaterer kontaktlisten."
    lNor.sCompSend0 = "Sett i kø"
    lNor.sCompSend0_ = "Send meldingen med prioritetsnivå 2."
    lNor.sCompSend1 = "Prioriter"
    lNor.sCompSend1_ = "Send meldingen med prioritetsnivå 1."
    lNor.sfrmAccessInf = "Nordicmafia pålogging"
    lNor.sfrmGenConf = "Generell oppsett"
    lNor.sfrmContList = "Kontaktliste"
    lNor.sinDelAll = "Slett ALLE meldinger"
    lNor.sinDelAll_ = "Tøm innboksen."
    lNor.sinDelThese = "Slett disse meldingene"
    lNor.sinDelThese_ = "Sletter de meldingene du ser i Postal's innboks nå."
    lNor.sinFilter0 = "Vis alle"
    lNor.sinFilter0_ = "Viser alle meldinger."
    lNor.sinFilter1 = "Vis uleste"
    lNor.sinFilter1_ = "Viser kun meldinger du ikke har lest enda."
    lNor.sinFilter2 = "Vis leste"
    lNor.sinFilter2_ = "Viser bare meldinger du alt har lest."
    lNor.sinStats = "%1 totalt, %2 nye"
    lNor.slbAppStatus = "Programstatus"
    lNor.slbAutocheck = "Hent PMs v/ oppstart"
    lNor.slbAutocheck_ = "Laster ned de 15. første PMene når du starter Postal. Anbefalt."
    lNor.slbAutosync = "Automatisk oppdatering"
    lNor.slbAutosync_ = "Ser etter nye PMs hvert x. sekund. Sett til 0 for å deaktivere funksjonen."
    lNor.slbCntLaunch = "Sjekk kontakter v/ start"
    lNor.slbCntLaunch_ = "Sjekker hvilke av dine kontakter som er påloggede når du starter postal."
    lNor.slbCntRefresh = "Autosjekk kontakter"
    lNor.slbCntRefresh_ = "Oppdaterer kontaktlisten hvert x. sekund. Sett til 0 for å deaktivere funksjonen."
    lNor.slbCntWarn = "Status-advarsel"
    lNor.slbCntWarn_ = "Advarer deg om en kontakt logger på eller av."
    lNor.slbInbox = ".:: Innboks ::."
    lNor.slbMaxread = "Les maks x meldinger"
    lNor.slbMaxread_ = "Høyeste antall meldinger du tillater Postal å lese om gangen. Sett til 0 for å lese alle 15."
    lNor.slbMsgComp = "Skriv en melding"
    lNor.slbMsgView = "Meldingsvisning"
    lNor.slbOutbox = ".:: Utboks ::."
    lNor.slbPassword = "Passord"
    lNor.slbSendCurr = "Send nåv. PM"
    lNor.slbSendCurr_ = "Om denne boksen er markert, blir alle kontaktene du klikker på tilsendt sist skrevet PM."
    lNor.slbUsername = "Brukernavn"
    lNor.soutStats = "%1 i kø. Neste: %2"
    lNor.sViewDel = "Slett"
    lNor.sViewDel_ = "Slett denne meldingen fra både Postal og NM. Om du lagret meldingen med ""Lagre"" knappen, vil det eksemplaret forbli beholdt."
    lNor.sViewReply = "Svar"
    lNor.sViewReply_ = "Klikk her for å skrive et svar til denne PMen."
    lNor.sViewSave = "Lagre"
    lNor.sViewSave_ = "Lagre denne meldingen til Postal's mappe."
    lNor.sViewUnread0 = "Ny melding"
    lNor.sViewUnread1 = "Gammel melding"
    
    lNor.sMsgDelThese = "Er du sikker på at du vil slette meldingene som vises i innboksen nå?"
    lNor.sMsgDelRead = "Er du sikker på at du vil slette alle leste meldinger?"
    lNor.sMsgDelAll = "Er du sikker på at du vil slette alle meldinger?"
    lNor.sMsgAddCnt = "Vennligst skriv NM-navnet på personen du vil legge til."
    lNor.sMsgRemCnt = "Er du sikker på at du vil fjerne ""%1"" fra din kontaktliste?"
    lNor.sMsgAdvSure = "Er du sikker på at du vil gjøre dette? Kontoen blir nærmest garantert modkillet."
    lNor.sMsgDelThisMsg = "Er du sikker på at du vil slette denne meldingen?"
End Sub

Sub Lang_Set(ByVal Lang As String)
    If Lang = "NOR" Then Lng = lNor
    If Lang = "ENG" Then Lng = lEng
    
    frmMain.lbInbox = Lng.slbInbox
    frmMain.lbOutbox = Lng.slbOutbox
    frmMain.lbAppStatus.Caption = Lng.slbAppStatus
    frmMain.frmAccessInf.Caption = Lng.sfrmAccessInf
    frmMain.frmGenConf.Caption = Lng.sfrmGenConf
    frmMain.frmContList.Caption = Lng.sfrmContList
    frmMain.lbUsername.Caption = Lng.slbUsername
    frmMain.lbPassword.Caption = Lng.slbPassword
    frmMain.accUser.ToolTipText = Lng.saccUser_
    frmMain.accPass.ToolTipText = Lng.saccPass_
    frmMain.accSave.Caption = Lng.saccSave
    frmMain.accSave.ToolTipText = Lng.saccSave_
    frmMain.lbMaxRead.Caption = Lng.slbMaxread
    frmMain.lbMaxRead.ToolTipText = Lng.slbMaxread_
    frmMain.lbAutosync.Caption = Lng.slbAutosync
    frmMain.lbAutosync.ToolTipText = Lng.slbAutosync_
    frmMain.lbAutocheck.Caption = Lng.slbAutocheck
    frmMain.lbAutocheck.ToolTipText = Lng.slbAutocheck_
    frmMain.lbCntRefresh.Caption = Lng.slbCntRefresh
    frmMain.lbCntRefresh.ToolTipText = Lng.slbCntRefresh_
    frmMain.lbCntWarn.Caption = Lng.slbCntWarn
    frmMain.lbCntWarn.ToolTipText = Lng.slbCntWarn_
    frmMain.lbCntLaunch.Caption = Lng.slbCntLaunch
    frmMain.lbCntLaunch.ToolTipText = Lng.slbCntLaunch_
    frmMain.cfgSave.Caption = Lng.scfgSave
    frmMain.cfgSave.ToolTipText = Lng.scfgSave_
    frmMain.in_Filter(0).Caption = Lng.sinFilter0
    frmMain.in_Filter(0).ToolTipText = Lng.sinFilter0_
    frmMain.in_Filter(1).Caption = Lng.sinFilter1
    frmMain.in_Filter(1).ToolTipText = Lng.sinFilter1_
    frmMain.in_Filter(2).Caption = Lng.sinFilter2
    frmMain.in_Filter(2).ToolTipText = Lng.sinFilter2_
    frmMain.in_DelThese.Caption = Lng.sinDelThese
    frmMain.in_DelThese.ToolTipText = Lng.sinDelThese_
    frmMain.in_DelAll.Caption = Lng.sinDelAll
    frmMain.in_DelAll.ToolTipText = Lng.sinDelAll_
    frmMain.cmdRefresh.Caption = Lng.scmdRefresh
    frmMain.cmdRefresh.ToolTipText = Lng.scmdRefresh_
    frmMain.cmdFilter.Caption = Lng.scmdFilter
    frmMain.cmdFilter.ToolTipText = Lng.scmdFilter_
    frmMain.cmdDeleteRead.Caption = Lng.scmdDeleteRead
    frmMain.cmdDeleteRead.ToolTipText = Lng.scmdDeleteRead_
    frmMain.cmdDelBy.Caption = Lng.scmdDelBy
    frmMain.cmdDelBy.ToolTipText = Lng.scmdDelBy_
    frmMain.cmdNew.Caption = Lng.scmdNew
    frmMain.cmdNew.ToolTipText = Lng.scmdNew_
    frmMain.cmdAdvertise.Caption = Lng.scmdAdvertise
    frmMain.cmdAdvertise.ToolTipText = Lng.scmdAdvertise_
    frmMain.cntAdd.Caption = Lng.scntAdd
    frmMain.cntAdd.ToolTipText = Lng.scntAdd_
    frmMain.cntRem.Caption = Lng.scntRem
    frmMain.cntRem.ToolTipText = Lng.scntRem_
    frmMain.cntSendPM.Caption = Lng.scntSendPM
    frmMain.cntSendPM.ToolTipText = Lng.scntSendPM_
    frmMain.cntRefresh.Caption = Lng.scntRefresh
    frmMain.cntRefresh.ToolTipText = Lng.scntRefresh_
    frmMain.lbSendCurr.Caption = Lng.slbSendCurr
    frmMain.lbSendCurr.ToolTipText = Lng.slbSendCurr_
    frmMain.lbMsgComp.Caption = Lng.slbMsgComp
    frmMain.Comp_Send(0).Caption = Lng.sCompSend0
    frmMain.Comp_Send(0).ToolTipText = Lng.sCompSend0_
    frmMain.Comp_Send(1).Caption = Lng.sCompSend1
    frmMain.Comp_Send(1).ToolTipText = Lng.sCompSend1_
    frmMain.lbMsgView.Caption = Lng.slbMsgView
    frmMain.View_Reply.Caption = Lng.sViewReply
    frmMain.View_Reply.ToolTipText = Lng.sViewReply_
    frmMain.View_Save.Caption = Lng.sViewSave
    frmMain.View_Save.ToolTipText = Lng.sViewSave_
    frmMain.View_Delete.Caption = Lng.sViewDel
    frmMain.View_Delete.ToolTipText = Lng.sViewDel_
End Sub
