; =======================================================================================
; ------------------------- OPEN SERVICE-NOW RECORDS ------------------------------------
; =======================================================================================
#SingleInstance Force

#NoEnv
SetWorkingDir, %A_ScriptDir%  ; Ensure script runs from its directory
companyFile := A_ScriptDir . "\CompanySettings.ini"
defaultCompany := "mycompany"
defaultTicketLength := 7  ; Default ServiceNow ticket number length

; Read company name, ticket length, and "Never Ask Again" setting from INI file
IniRead, storedCompany, %companyFile%, Settings, CompanyName, %defaultCompany%
IniRead, neverAskAgain, %companyFile%, Settings, NeverAskAgain, 0  ; Default is '0' (ask again)
IniRead, ticketLength, %companyFile%, Settings, TicketLength, %defaultTicketLength%

; Ensure ticketLength falls within the valid range (7-10)
if (ticketLength < 7 or ticketLength > 10) {
    ticketLength := defaultTicketLength
    IniWrite, %ticketLength%, %companyFile%, Settings, TicketLength
}

; Create Tray Menu Option for Reset
Menu, Tray, Add, Reset Settings, ResetConfig
Menu, Tray, Add, Exit, ExitScript

; If "Never Ask Again" is enabled, set company name and skip GUI
If (neverAskAgain = 1) {
    global companyName := storedCompany
    Return
}

Gui, Add, Text, x10 y10 w300 h25, **Customize Your ServiceNow Settings**
Gui, Add, Text, x10 y30 w400 h20, ______________________________________________________________________________
Gui, Font, S10 Bold
Gui, Add, Text, x10 y50 w400 h25, Enter the company name as used in the ServiceNow URL:
Gui, Font, Norm
Gui, Add, Edit, x10 y80 w250 h25 vCompanyInput, %storedCompany%
Gui, Add, Text, x10 y110 w300 h25, Example: `https://<company>.service-now.com/`
Gui, Add, Text, x10 y130 w400 h20, ______________________________________________________________________________

Gui, Font, S10 Bold
Gui, Add, Text, x10 y150 w350 h25, Select the number of digits used in ticket numbers:
Gui, Font, Norm
Gui, Add, Edit, x10 y180 w50 h25 vTicketLengthInput, %ticketLength%
Gui, Add, Text, x10 y210 w300 h25, Example: `INC#######` (7 digits is default)
Gui, Add, Text, x10 y230 w400 h20, ______________________________________________________________________________

Gui, Font, S10 Bold
Gui, Add, Text, x10 y250 w300 h25, Select this to remove this pop-up at startup:
Gui, Font, Norm
Gui, Add, Checkbox, x10 y275 w300 h25 vNeverAskAgain, Never ask again
Gui, Add, Text, x10 y295 w400 h20, ______________________________________________________________________________
Gui, Add, Text, x10 y315 w450 h65, You can clear the config by right clicking the icon in the system tray and selecting "Reset Settings"

Gui, Add, Button, x10 y360 w100 h30 gSaveCompany, Save
Gui, Add, Button, x120 y360 w100 h30 gExitGui, Cancel
Gui, Show, w450 h400, Company Setup


Return  ; End of Auto-executing Section

; Save Button Action
SaveCompany:
    Gui, Submit, NoHide  ; Retrieve input

    ; Validate ticket length input
    if (TicketLengthInput < 7 or TicketLengthInput > 10) {
        MsgBox, Invalid ticket number length! Must be between 7-10.
        Return
    }

    ; Store values in the INI file
    IniWrite, %CompanyInput%, %companyFile%, Settings, CompanyName
    IniWrite, %NeverAskAgain%, %companyFile%, Settings, NeverAskAgain
    IniWrite, %TicketLengthInput%, %companyFile%, Settings, TicketLength

    if (NeverAskAgain = 0)
        AskStatus := "You will be asked again"
    else if (NeverAskAgain = 1)
        AskStatus := "You will no longer be asked to update the config file."

    
    global companyName := CompanyInput
    global ticketLength := TicketLengthInput
    Gui, Destroy
    MsgBox, Company name saved as "%CompanyInput%"! `n`nTicket number length: %TicketLengthInput% `n`nConfig Status: %AskStatus% 
Return

; Cancel Button Action
ExitGui:
    Gui, Destroy
Return

; Reset Config (Triggered from Tray Menu)
ResetConfig:
    IniDelete, %companyFile%, Settings, CompanyName
    IniDelete, %companyFile%, Settings, NeverAskAgain
    IniDelete, %companyFile%, Settings, TicketLength
    MsgBox, Configuration reset! Relaunch the script to enter new values.
Return


; Cancel Button Action
ExitScript:
    Gui, Destroy
Return

^+O::OpenTICKET("OpenTICKET")  ; Hotkey: Ctrl + Shift + O triggers OpenTICKET function.


OpenTICKET(destination) {
    ;#warn
    SetWorkingDir %A_ScriptDir%  ; Ensures script operates from its directory.
    

    Gui,2: New
    Gui,2: default
    Gui,2: -dpiscale   ; Disables DPI scaling (could be adjusted for high-DPI screens).
    GUI,2: Font,cGray,Fixedsys  ; Sets font color and type.

    ; Creates UI buttons
    Gui,2: Add, Button,  x25 y10 w150 h25 gBT1, INC Ticket  
    Gui,2: Add, Button,  x25 y40 w150 h25 gBT2, CTASK
    Gui,2: Add, Button,  x25 y70 w150 h25 gBT3, CHG Change
    Gui,2: Add, Button,  x25 y100 w150 h25 gBT4, SCTASK
    Gui,2: Add, Button,  x25 y130 w150 h25 gBT5, REQ
    Gui,2: Add, Button,  x25 y160 w150 h25 gBT6, RITM
    Gui,2: Add, Button,  x25 y190 w150 h25 gBT7, TASK
    Gui,2: Add, Button,  x25 y220 w150 h25 gBT8, KB
    Gui,2: Add, Button,  x25 y250 w150 h25 gBT9, User Lookup
    Gui,2: Add, Button,  x25 y280 w150 h25 gBT10, Group Lookup
    Gui,2: Add, Button,  x25 y310 w150 h25 gBT11, CI Lookup
    
    Gui,2: Show, x100 y10 w200 h350, SNow  ; Displays GUI
    
    send, {down}{up}  ; Quick keystroke to ensure focus is set
    
    If ErrorLevel {  
        MsgBox, Error launching ticket interface!
        Exit
    }  
    
    Return
}

; ---------------------- UNIVERSAL FUNCTION TO HANDLE REQUESTS ----------------------


OpenRecord(type) {
    global ticketLength  ; Ensure ticket length is accessible

    ; Adjust regex pattern dynamically based on ticketLength
    regexPattern := "\b" type "\d{" ticketLength "}\b"

    RegExMatch(Clipboard, regexPattern, RecordID)

    InputBox, userInput, Open Service-Now Record, Enter a %type% number (full or short).`r`nFull numbers are populated from the clipboard.,,,,,,,,%RecordID%

    If ErrorLevel {
        MsgBox, Operation canceled.
        Exit
    }

    RegExMatch(userInput, regexPattern, RecordID)
    If (!RecordID) {
        RegExMatch(userInput, "\b\d{1," ticketLength "}\b", RecordID)
        If (RecordID) {
            StringLen, length, RecordID
            StringLeft, start, RecordID, 1
            If !(length = ticketLength and start >= 4) {
                length := ticketLength - length
                Loop %length% {
                    RecordID := "0" . RecordID
                }
                RecordID := type . RecordID
            }
        }
    }


    recordURLs := Object()
    recordURLs["INC"] := "incident.do?sysparm_query=number="
    recordURLs["CHG"] := "change_request.do?sysparm_query=number="
    recordURLs["CTASK"] := "change_task.do?sysparm_query=number="
    recordURLs["SCTASK"] := "sc_task.do?sysparm_query=number="
    recordURLs["REQ"] := "sc_request.do?sysparm_query=number="
    recordURLs["RITM"] := "sc_req_item.do?sysparm_query=number="
    recordURLs["TASK"] := "task.do?sysparm_query=number="
    recordURLs["KB"] := "kb_knowledge.do?sysparm_query=number="
    

    recordType := type  ; Store the record type
    urlPath := recordURLs[recordType]  ; Retrieve the correct URL path
    fullURL := "https://" . companyName . ".service-now.com/nav_to.do?uri=" urlPath RecordID
    ;MsgBox, %fullURL%  ; Show the constructed URL for debugging
    Run, %fullURL%  ; Open the record in a browser
    
    

    Gosub, LabelExit  
    Return
}

; ---------------------- LOOKUP FUNCTIONS ----------------------

OpenLookup(type) {
    InputBox, userInput, Lookup %type% in Service-Now, Enter %type% information (e.g.`, email for users`, name for groups/CI).,,,

    If ErrorLevel {
        MsgBox, Operation canceled.
        Exit
    }

    lookupURLs := Object()
    lookupURLs["User"] := "sys_user.do?sysparm_query=email="
    lookupURLs["Group"] := "sys_user_group.do?sysparm_query=name="
    lookupURLs["CI"] := "cmdb_ci.do?sysparm_query=name="

    recordType := type  ; Store the type in a variable
    urlPath := recordURLs[recordType]  ; Retrieve the corresponding URL path
    fullURL := "https://" . companyName . ".service-now.com/nav_to.do?uri=" . urlPath . userInput
    Run, %fullURL%

    Gosub, LabelExit  
    Return
}

; ---------------------- BUTTON ACTIONS ----------------------

bt1:
Gosub, OpenRecordINC
Return

bt2:
Gosub, OpenRecordCTASK
Return

bt3:
Gosub, OpenRecordCHG
Return

bt4:
Gosub, OpenRecordSCTASK
Return

bt5:
Gosub, OpenRecordREQ
Return

bt6:
Gosub, OpenRecordRITM
Return

bt7:
Gosub, OpenRecordTASK
Return

bt8:
Gosub, OpenRecordKB
Return

bt9:
Gosub, OpenLookupUser
Return

bt10:
Gosub, OpenLookupGroup
Return

bt11:
Gosub, OpenLookupCI
Return

OpenRecordINC:
OpenRecord("INC")
Return

OpenRecordCTASK:
OpenRecord("CTASK")
Return

OpenRecordCHG:
OpenRecord("CHG")
Return

OpenRecordSCTASK:
OpenRecord("SCTASK")
Return

OpenRecordREQ:
OpenRecord("REQ")
Return

OpenRecordRITM:
OpenRecord("RITM")
Return

OpenRecordTASK:
OpenRecord("TASK")
Return

OpenRecordKB:
OpenRecord("KB")
Return

OpenLookupUser:
OpenLookup("User")
Return

OpenLookupGroup:
OpenLookup("Group")
Return

OpenLookupCI:
OpenLookup("CI")
Return

LabelExit:
Gui, Destroy  ; Closes the GUI window
Exit
Return