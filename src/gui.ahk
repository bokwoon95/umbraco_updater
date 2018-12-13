#Persistent
#SingleInstance Force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

RegRead, un, HKEY_CURRENT_USER\SOFTWARE\Update_RMC_Companies, un
RegRead, pw, HKEY_CURRENT_USER\SOFTWARE\Update_RMC_Companies, pw
Gui, Add, Button, x322 y29 w-310 h-90 , Button
Gui, Add, Text, x22 y9 w180 h20 , Umbraco username
Gui, Add, Edit, x22 y29 w180 h20 vUsername, %un%
Gui, Add, Text, x22 y59 w180 h20 , Umbraco password
Gui, Add, Edit, x22 y79 w180 h20 Password vPassword, %pw%
Gui, Add, Button, x22 y119 w180 h60 gPickXlsx, Choose an excel/html file to upload
; Generated using SmartGUI Creator 4.0
Gui, Show, x165 y136 h208 w229, Admaterials
Return

PickXlsx:
    GuiControlGet, Username
    GuiControlGet, Password

    if (Username == "" || Password == "") {
        MsgBox Both username and password must be filled in
        return
    }
    if (Username != un || Password != pw) {
        RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Update_RMC_Companies, un, %Username%
        RegWrite, REG_SZ, HKEY_CURRENT_USER\SOFTWARE\Update_RMC_Companies, pw, %Password%
    }

    FileSelectFile, RMC_certifications, 1,,,*.xlsx;*.xls;*.xlsm;*xlst;*.html
    if (RMC_certifications == "") {
        return
    }
    FoundPos := RegexMatch(RMC_certifications, "^(?P<Filename>.*)\.(?P<Extension>.*)$", Match)
    if (RegexMatch(MatchExtension, "^xls.?$")) {
        SplashTextOn, 600,100,Converting excel document to html,Please wait, converting excel document to html..`n`n(PRESS ESC TO ABORT ANYTIME)
        WinMove, Converting excel document to html,,0,0
        #Include excel2html.ahk
        Extract_excel2html(A_ScriptDir . "\excel2html.exe")
        RunWait, %comspec% /c %A_ScriptDir%\excel2html.exe "%RMC_certifications%"
        RMC_certifications_html := MatchFilename . ".html"
        FileDelete %A_ScriptDir%\excel2html.exe
        SplashTextOff
    } else if (RegexMatch(MatchExtension, "^html$")) {
        RMC_certifications_html := RMC_certifications
    } else {
        MsgBox the %MatchFilename%.html file was not generated
        return
    }

    SplashTextOn, 600,100,Opening Chrome,Please wait, opening Google Chrome`n`n(PRESS ESC TO ABORT ANYTIME)
    WinMove, Opening Chrome,,0,0
    ; Create Chrome Instance
    driver:= ComObjCreate("Selenium.CHROMEDriver")
    driver.Get("http://www.admaterials.com.sg/umbraco")
    driver.Window.Maximize()
    driver.WaitForScript("document.querySelector('#login > div > div:nth-child(2) > form > div:nth-child(1) > input')", ,10000)
    driver.Wait(500)
    SplashTextOff

    ; Enter username and password
    field := driver.FindElementByXPath("//*[@id='login']/div/div[1]/form/div[1]/input")
    field.SendKeys(Username)
    field := driver.FindElementByXPath("//*[@id='login']/div/div[1]/form/div[2]/input")
    field.SendKeys(Password)
    field.Submit()
    SplashTextOn, 600,100,Edit page,Navigating to page editor..`n`n(PRESS ESC TO ABORT ANYTIME)
    WinMove, Edit page,,0,0
    driver.Wait(1000)
    try {
        error := driver.FindElementByXPath("//*[@id='login']/div/div[1]/form/div[3]/div")
        if (error) {
            MsgBox Login failed for user %Username%`nPlease check your username and password
            driver :=
            return
        }
    }

    ; Close any update prompts (that may block the block the '<>' source code button from being clicked)
    try {
        button := driver.FindElementByXPath("//*[@id='umb-notifications-wrapper']/ul/li/a")
        button.Click()
    }
    ; #umb-notifications-wrapper > ul > li > a
    ; //*[@id='umb-notifications-wrapper']/ul/li/a

    ; View table's source code
    ; Wait until the '<>' source code button renders, then click on it
    driver.Get("http://www.admaterials.com.sg/umbraco#/content/content/edit/1250")
    SplashTextOff
    SplashTextOn, 600,100,Editing table,Waiting for source code button '<>' to render..`n`n(PRESS ESC TO ABORT ANYTIME)
    WinMove, Editing table,,0,0
    driver.WaitForScript("document.querySelector('#mceu_0 > button')", ,300000) ; 5 minute timeout
    driver.Wait(1000)
    button := driver.FindElementByXPath("//*[@id='mceu_0']/button")
    button.Click()
    SplashTextOff

    ; Copy html file contents into clipboard
    tempvar := ClipboardAll
    FileRead, Clipboard, %RMC_certifications_html%
    ; Convert any unicode characters into their html equivalent codes
    Clipboard := RegexReplace(Clipboard, "½", "&frac12;")
    Clipboard := RegexReplace(Clipboard, "“", "&ldquo;")
    Clipboard := RegexReplace(Clipboard, "”", "&rdquo;")

    ; Paste clipboard contents into textarea
    textarea := driver.FindElementByXPath("//*[@id='mceu_39']")
    textarea.Clear()
    textarea.SendKeys(driver.Keys.Control, "v")
    Clipboard := tempvar
    tempvar :=

    ; Click OK and preview the changes
    button := driver.FindElementByXPath("//*[@id='mceu_41']/button")
    button.Click()
    button := driver.FindElementByXPath("//*[@id='contentcolumn']/div/div/form/div/div[3]/div/div[2]/div[1]/button")
    button.Click()
return


GuiClose:
ExitApp

Esc::ExitApp
