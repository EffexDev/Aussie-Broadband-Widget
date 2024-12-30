#Requires AutoHotkey v2.0
#Include Templates.ahk

; ALt-F1 to entirely reload the script. This is mostly for testing and debugging.
!F1::
{
    Reload
}

IniWrite("xxx", "config.ini", "Customer", "CustomerName")

Global LiveChatMap := Map()
Global GeneralNBNMap := Map()
Global GeneralHardwareMap := Map()
Global GeneralOTRSMap := Map()
Global RedmineJiraMap := Map()
Global PPMap := Map()
Global FHMap := Map()
Global ReconnectionMap := Map()
Global CallbackMap := Map()
Global FaultTemplatesMap := Map()
Global DiscoveryMap := Map()
Global SpeedsMap := Map()
Global DropoutsMap := Map()
Global ConnectionMap := Map()
Global SetupMap := Map()
Global LinkMap := Map()
Global HardwareMap := Map()
Global ServiceRequestMap := Map()
Global DelaysMap := Map()
Global HFCMap := Map()
Global FTTCMap := Map()
Global FTTPMap := Map()
Global ValidationMap := Map()
Global BanlistingMap := Map()
Global FailedPaymentMap := Map()
Global PaymentMap := Map()
Global NBNCompMap := Map()
Global RaisingMap := Map()
Global ClarificationMap := Map()
Global ResolutionsMap := Map()
Global ChangesMap := Map()
Global TIOMap := Map()
Global ContactsMap := Map()
Global TCSGeneralMap := Map()
Global TCSBillingMap := Map()
Global TCSSuspensionMap := Map()
Global TCSChangesMap := Map()

; Declarations of global variables. This is used to make the popout visible to other functions
Global NotesGui := Gui(,"Notepad"), Notes := NotesGui.Add("Edit", "h600 w685", "")
Global TemplatesGui := Gui(,"Templates"), Templates := TemplatesGui.Add("Edit", "h600 w685", "")
Global HotkeysGui := Gui(,"Useful Links"), Hotkeys := HotkeysGui.Add("Text", "+Wrap h230 cFFFFFF", "Available Hotkeys:`n`n@@ - Your email`n`n~~ - Date 7 calendar days from now. For abandoment comms`n`nCTRL+Shift+S - Randomized signature for app faults. Also checks the publish to app checkbox`n`nCTRL+DEL - Content aware search ie. Superlookup")

Hotkeys.SetFont("s10","Nunito")
Notes.SetFont("s10","Nunito")
Templates.SetFont("s10","Nunito")

NotesGui.BackColor := "c005900"
TemplatesGui.BackColor := "c005900"
HotkeysGui.BackColor := "c005900"

; The main body of the GUI itself. Dimensions and tabs etc
Global AussieGui := Gui("-Caption +Border","Aussie Tool Kit V2.0")
AussieGui.BackColor := "c005900"
AussieGui.Add("Picture", "ym+10 x+20 w215 h-1","AussieLogo.png")


AussieGui.SetFont("s10 c000000","Nunito")
AussieGui.Add("Text", " xm cFFFFFF" , "Customer Name:")
Global CustomerNameField := AussieGui.Add("Edit", "yp-3 xm+105 w150 vCustomerNameValue", "").OnEvent("Change", CustomerNameEdit)
TemplateTab := AussieGui.Add("Tab3","xm h70 w610 BackgroundWhite", ["General", "Accounts", "Faults","Order Support","Complaints","T and Cs"])
ToolsTab := AussieGui.Add("Tab3", "WP h80 BackgroundWhite c222222 vToolsTab", ["QOL", "Automations", "Options"])

AussieGui.Show("x1920 y0 w630")

;DO NOT REMOVE
Send "xxx"

;First set of tabs, for department selection to segregate templates and keep things organised. This grabs the options selected in both dropdowns and saves them into a variable to be used later.
TemplateTab.UseTab(1)
SelGeneralReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedGeneralReason Choose1", GeneralReasons)
SelGeneralReason.OnEvent('Change', SelGeneralReasonSelected)
SelGeneralTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedGeneral", GeneralTemplates
[SelGeneralReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunGeneral)

TemplateTab.UseTab(2)
SelAccountReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedAccountReason Choose1", AccountReasons)
SelAccountReason.OnEvent('Change', SelAccountReasonSelected)
SelAccountTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedAccount", AccountTemplates
[SelAccountReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunAccount)

TemplateTab.UseTab(3)
SelFaultReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedFaultReason Choose1", FaultReasons)
SelFaultReason.OnEvent('Change', SelFaultReasonSelected)
SelFaultTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedFault", FaultTemplates[SelFaultReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunFault)

TemplateTab.UseTab(4)
SelDeliveryReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedDeliveryReason Choose1", DeliveryReasons)
SelDeliveryReason.OnEvent('Change', SelDeliveryReasonSelected)
SelDeliveryTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedDelivery", DeliveryTemplates[SelDeliveryReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunDelivery)

TemplateTab.UseTab(5)
SelComplaintReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedComplaintReason Choose1", ComplaintReasons)
SelComplaintReason.OnEvent('Change', SelComplaintReasonSelected)
SelComplaintTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedComplaint", ComplaintTemplates[SelComplaintReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunComplaint)

TemplateTab.UseTab(6)
SelTCSReason := AussieGui.AddDropDownList("w200 h100 r20 BackgroundFFFFFF vPickedTCSReason Choose1", TCSReasons)
SelTCSReason.OnEvent('Change', SelTCSReasonSelected)
SelTCSTemplate := AussieGui.AddDropDownList("yp w200 r20 BackgroundFFFFFF vPickedTCS", TCSTemplates[SelTCSReason.Value])
GenerateFault := AussieGui.Add("Button", "yp", "Generate").OnEvent("Click", RunTCS)

; The tools tab. Controls the bottom set of tabs and the content of them. Functions are below
ToolsTab.UseTab(1)
Superlookup := AussieGui.Add("Button","xm+15 y245 vSuperlookup", "Superlookup").OnEvent("Click", ProcessSuperlookup)
AussieGui.Add("Button", "yp w90", "Notes").OnEvent("Click", NotePad)
AussieGui.Add("Button","yp w90", "Hotkeys").OnEvent("Click", HotkeysPad)
AussieGui.Add("Button","yp w90","Startup").OnEvent("Click", Startup)
AussieGui.Add("Button","yp w100", "Lock Terminal").OnEvent("Click", LockTerminal)
NotePadEmbedded := AussieGui.Add("Edit", "yp+40 xm+10 h510 w585 vNotePadEmbedded", "")

ToolsTab.UseTab(2)
AussieGui.Add("Button", "xm+15 y245 Section", "Ping Test").OnEvent("Click", PingTest)
AussieGui.Add("Button", "yp", "Traceroute").OnEvent("Click", Traceroute)
AussieGui.Add("Button", "yp", "NSLookup").OnEvent("Click", NSLookup)
AussieGui.Add("Button", "yp", "Prorata Calc").OnEvent("Click", ProRataCalc)

ToolsTab.UseTab(3)
Global AlwaysOnTopButton := AussieGui.Add("Checkbox", "xm+15 y245 Section vAlwaysOnTop ").OnEvent("Click", AlwaysOnTopToggle)
AlwaysOnTopCheckBoxText := AussieGui.Add("Text", "yp xp+20 c000000", "Always on Top")
Global ShowNotesButton := AussieGui.Add("Checkbox", "yp x+20 vShowNotesButton").OnEvent("Click", ShowNotes)
ShowNotesButtonText := AussieGui.Add("Text", "yp xp+20 c000000", "Show Notepad")
Global DarkmodeButton := AussieGui.Add("Checkbox", "yp x+20 vDarkModeButton ").OnEvent("Click", Darkmode)
DarkmodeButtonText := AussieGui.Add("Text", "yp xp+20 c000000", "Darkmode")
Global UpdateButton := AussieGui.Add("Button", "yp-5 x+120", "Check for Updates").OnEvent("Click", UpdateWidgetCheck)
AussieGui["NotePadEmbedded"].Visible := 0


; This section controls cascading dropdowns. It will clear the second dropdown field and replace it when the first is altered. This allows you to have multiple categories per department.
global CustomerName := ""  ; Initialize at script level

CustomerNameEdit(CustomerNameValue, *) {
    global CustomerName
    CustomerName := AussieGui["CustomerNameValue"].Value
    global CustomerNameSanitised := RegExReplace(CustomerName, "[^a-zA-Z]", "")
    IniWrite(CustomerNameSanitised, "config.ini", "Customer", "CustomerName")
    UpdateTemplates()
}

SelGeneralReasonSelected(*) 
{
    SelGeneralTemplate.Delete()
    SelGeneralTemplate.Add(GeneralTemplates[SelGeneralReason.value])
    SelGeneralTemplate.Choose(1)
}

SelAccountReasonSelected(*) 
{
    SelAccountTemplate.Delete()
    SelAccountTemplate.Add(AccountTemplates[SelAccountReason.value])
    SelAccountTemplate.Choose(1)
}

SelFaultReasonSelected(*) 
{
    SelFaultTemplate.Delete()
    SelFaultTemplate.Add(FaultTemplates[SelFaultReason.value])
    SelFaultTemplate.Choose(1)
}

SelDeliveryReasonSelected(*) 
{
    SelDeliveryTemplate.Delete()
    SelDeliveryTemplate.Add(DeliveryTemplates[SelDeliveryReason.value])
    SelDeliveryTemplate.Choose(1)
}

SelComplaintReasonSelected(*) 
{
    SelComplaintTemplate.Delete()
    SelComplaintTemplate.Add(ComplaintTemplates[SelComplaintReason.value])
    SelComplaintTemplate.Choose(1)
}

SelTCSReasonSelected(*) 
{
    SelTCSTemplate.Delete()
    SelTCSTemplate.Add(TCSTemplates[SelTCSReason.value])
    SelTCSTemplate.Choose(1)
}

#Include Settings.ahk
#Include Generate.ahk
#Include FunctionLibrary.ahk
