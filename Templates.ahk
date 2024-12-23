#SingleInstance Force

; This is mostly just a bunch of text so I am not going to comment the whole lot of it. 
; 
; The basic gist is that each department tab has an array for the first dropdown and an array of arrays for the second dropdown.
; Each item in the first dropdowns arrays has an associated map containing key-value pairs. The keys in these pairs correlate with the items in the associated template array.
;
; The main script uses the two selected dropdown options to pull the value from the key-value pair and output it to a text field.

global CustomerName := IniRead("config.ini", "Customer", "CustomerName")

csFullName := A_UserName
csFirstName := ""

if (RegExMatch(csFullName, "^[^.]*",&csFirstName)) {
csTitle:=StrTitle(csFirstName[0])
}

CurrentDate := FormatTime(, "yyyyMMdd")
NewDate := FormatTime(DateAdd(CurrentDate, 7, "days"), "dd/MM/yyyy")

GeneralReasons := ["LiveChat","NBN","General Hardware","OTRS", "Redmine/Jira"]
GeneralTemplates := [["Direct Link"],["Raise"],["Resets", "Router Swap"],["Charges"],["Redmine Bug", "Redmine Enhancement"]]

AccountReasons := ["Payment Plan", "Financial Hardship"]
AccountTemplates := [["Set Up","Confirmation", "Payment Options", "Update Details"],["Template"]]

FaultReasons := ["General Templates","Discovery","Slow Speeds", "Dropouts", "No Connection", "Service Setup", "CSP Linking", "Hardware"]
FaultTemplates := [["Titles", "Warning", "Closing", "SRA/SOS/ROC"],["General Issue", "Slow Speeds", "Dropouts", "No Connection"],["Direct Connection", "Wifi Limitations", "Bandsteering", "Line Rates"],["Bandsteering", "Stability Profile", "SRA", "Cabling"],["UNI-D Port", "Outage"],["VDSL", "EWAN"],["NBN", "Tracking"],["Local Issue", "Cabling"]]

DeliveryReasons := ["HFC", "FTTP", "FTTC", "Validation", "Banlisting", "Missing Payment Info", "Service Request", "Delays"]
DeliveryTemplates := [["Day 1","Day 2", "Day 3", "Multiple NTDs","Signal Issue"],["Multiple UNI-D"],["Day 1","Day 2", "Day 3"],["Day 1","Day 2", "Day 3"],["Day 1"],["Day 1","Day 2", "Day 3"],["NTD/Address Mismatch", "FTTP - Rebuild", "FTTN/B - SC12 No CPI", "FTTN/B - SC13 No CPI", "FTTN/B - FNN/ULL Dispute", "FTTN/B - Rebuild", "FTTN/B - POTS Rejection", "HFC - Urgent Completion", "HFC - Remote Activation", "HFC - Multiple NTDs", "Order Escalation"], ["TC4", "In-Flight Order"]]

ComplaintReasons := ["NBN Complaints", "Raising", "Clarification", "Resolutions", "State Changes", "TIO", "Contacts"]
ComplaintTemplates := [["Templates"],["Templates"],["Templates"],["Templates"],["Changing PRD"],["Templates"],["SMS", "Email"]]

TCSReasons := ["General","Billing", "Suspension/Termination", "Changes"]
TCSTemplates := [["Availability", "Maintenance", "Support", "EU Obligations", "Data", "Access", ],["Fees", "Payment Methods", "Disputes","Other Fees"],["Service Use", "Poor Behaviour", "Payment Issues"],["TOS Changes"]]

;---------------- General -----------------

UpdateTemplates() {

    global TemplatesMaps := Map(
        "LiveChat", LiveChatMap,
        "NBN", GeneralNBNMap,
        "General Hardware", GeneralHardwareMap,
        "OTRS", GeneralOTRSMap,
        "Redmine/Jira", RedmineJiraMap,
        "Payment Plan", PPMap,
        "Financial Hardship", FHMap,
        "General Templates",FaultTemplatesMap,
        "Discovery", DiscoveryMap,
        "Slow Speeds", SpeedsMap,
        "Dropouts", DropoutsMap,
        "No Connection", ConnectionMap,
        "Service Setup", SetupMap,
        "CSP Linking", LinkMap,
        "Hardware", HardwareMap,
        "HFC", HFCMap,
        "FTTP", FTTPMap,
        "FTTC", FTTCMap,
        "Validation", ValidationMap,
        "Banlisting", BanlistingMap,
        "Missing Payment Info", PaymentMap,
        "Service Request", ServiceRequestMap,
        "Delays", DelaysMap,
        "NBN Complaints", NBNCompMap,
        "Raising", RaisingMap,
        "Clarification", ClarificationMap,
        "Resolutions", ResolutionsMap,
        "State Changes", ChangesMap,
        "TIO", TIOMap,
        "Contacts", ContactsMap,
        "General", TCSGeneralMap,
        "Billing", TCSBillingMap,
        "Suspension/Termination", TCSSuspensionMap,
        "Changes", TCSChangesMap
    )
    
    global CustomerName, csTitle

    Global LiveChatMap := Map(
        "Direct Link", ""
    ) 

    Global GeneralNBNMap := Map(
        "Raise", "",
    ) 

    Global GeneralHardwareMap := Map(
        "Resets", "",

        "Router Swap", "",
    )

    Global GeneralOTRSMap := Map(
        "Charges", "",
    )

    Global RedmineJiraMap := Map(
        "Redmine Bug", "",

        "Redmine Enhancement", ""
    )

    ; --------------- Accounts ----------------

    Global PPMap := Map(
        "Set Up", "",

        "Confirmation", "",

        "Payment Options", "",

        "Update Details", "",
    )

    Global FHMap := Map(
        "Template", "",
    )

    ;------------------- Faults -----------------------

    Global FaultTemplatesMap := Map(
        "Titles", "",

        "Warning", "",

        "Closing", "",

        "SRA/SOS/ROC", ""
    )

    Global DiscoveryMap := Map(
        "General Issue", "",

        "Slow Speeds", "",

        "Dropouts", "",
    )

    Global SpeedsMap := Map(
        "Direct Connection", "",

        "Wifi Limitations", "",

        "Bandsteering", "",

        "Line Rates", ""
    )

    Global DropoutsMap := Map(
        "Bandsteering", "",

        "Stability Profile", "",

        "SRA", "",

        "Cabling", ""
    )

    Global ConnectionMap := Map(
        "UNI-D Port", "",

        "Outage", ""
    )

    Global SetupMap := Map(
        "VDSL", "",

        "EWAN", ""
    )

    Global LinkMap := Map(
        "NBN", "NBN Appointment ID:`n`nNBN Incident ID:",
        "Tracking", "Return Post Tracking Code:"
    )

    Global HardwareMap := Map(
        "Local Issue", "",

        "Cabling", ""
    )

    Global ServiceRequestMap := Map(
        "NTD/Address Mismatch", "Request Type:  Equipment`nSeverity:  NTD Issue`n`nHi Team,`n`nYour records show that [INSERT NTD SERIAL NUMBER] is installed at [INSERT LOCID]. We have confirmed with our EU that actually [INSERT NTD SERIAL NUMBER] is installed onsite.`n`n- [INSERT PRI] has been deleted`n- Proof of Occupancy and Photos of the NTD (including Mac and Serial details) have been emailed to: proofofaddress@nbnco.com.au referencing this incident ID.`n`nPlease amend your records to reflect the correct NTD installed at [INSERT LOCID] so we can proceed with a connection for our EU.`n`nThank you.",

        "FTTP - Rebuild", "Hi Team,`n`nWe have been informed by our EU that [INSERT LOCID] has been knocked down and rebuilt and the cable and PCD have been removed from site. Please arrange to have [INSERT LOCID] rolled back to [INSERT ANTICIPATED SQ].`n`nPlease see the following supporting information:`n- Was the PCD previously installed onsite? YES`n- Can the EU locate where the PCD was located? YES/NO`n- Is the cabling from the PCD to the NTD original location still present? YES/NO`n- Who was responsible for the removal of this equipment?`nName`nMobile`nEmail`n`nPlease note Proof of Occupancy and supporting photos have been emailed to proofofaddress@nbnco.com.au referencing this incident number.`n`nPlease proceed with rolling back [INSERT LOCID] to [INSERT ANTICIPATED SQ] so we can arrange a connection for our EU.`n`nThank you",

        "FTTN/B - SC12 No CPI", "Request Type:  FNN/ULL Dispute`nSeverity: Low`nFNN: 0399999999`n`nHi Team,`n`nLOC is currently SC12 with no CPIs listed.`n`nPlease amend your records to reflect the SC12 CPI at this address or if there are no records of CPIs at this address, please roll back LOC to SC11. `n`nPlease rectify this issue so we can proceed with an order for our EU.`n`nThank you.",

        "FTTN/B - SC13 No CPI", "Request Type:  FNN/ULL Dispute`nSeverity: Low`nFNN: 0399999999`n`nHi Team,`n`nPlease amend your records to reflect the active SC13 CPI at [INSERT LOCID].`n`n[INSERT LOCID] is SC13 with no active CPI's and we can confirm that there has been an active nbn service at this address.`n`nPlease rectify this so we can raise an order for our EU.`n`nThank you.",

        "FTTN/B - FNN/ULL Dispute", "Request Type:  FNN/ULL Dispute`nSeverity:  Low`n`nHi Team`n`nPlease align FNN/ULL [INSERT FNN/ULL] against [INSERT LOCID] we can confirm that this is a valid match against this address.`n`nPlease rectify this so we can raise an order for our EU.`n`nThank you ",

        "FTTN/B - Rebuild", "Request Type:  Address`n`nSeverity:  Incorrect Service Class Returned`n`nHi Team,`n`nWe have been informed by our EU that [INSERT LOCID] has been knocked down and rebuilt and cabling has been removed from site. Please arrange to have [INSERT LOCID] rolled back to [INSERT ANTICIPATED SQ].`n`nPlease see the following supporting information:`n`nWas the NBN previously installed onsite? YES/NO`nIs the conduit from the Property boundary to the premise present? YES/NO is the internal cabling in place ? ( first socket)`nWho knocked it down/removed the cables?`nName`nEmail`nMobile`n`nPlease note Proof of Occupancy and supporting photos have been emailed to proofofaddress@nbnco.com.au referencing this incident number.`n`nPlease proceed with rolling back [INSERT LOCID] to [INSERT ANTICIPATED SQ] so we can arrange a connection for our EU.`n`nThank you. ",

        "FTTN/B - POTS Rejection", "Request Type:  Order`nSeverity:  Order Error`n`nHi Team,`n`nWe raised [ORDER ID] and this was rejected due to the following reason:`n`nOrder status: Rejected`nReason code: RJ005177`nDescription: The supplied FNN or ULL ID provided for the POTS Interconnect was not found against the Copper Pair identified in the request`n`nWe are trying to place our order against the inactive CPI at [INSERT LOCID] therefore a FNN/ULL is not required. Please amend your records so we can proceed with an order for our EU.`n`nThank you.",

        "HFC - Urgent Completion", "Request Type:  Equipment`nIssue:  Remote Activation Request`n`nHi Team,`n`nWe require [INSERT ORDER ID] to be completed urgently.`n`nThe EU has DHCP logs and recent usage against:`nLocation ID:`nAVC:`nC-TAG:`nProduct Instance ID:`n`nArris NTD details are as follows:`n[MAC ADDRESS]`n[SERIAL NUMBER]`nArris NTD light status: ON/OFF`n* 4 green light on - Y / N`n* Upstream light is flashing orange - Y / N`n* Downstream light is flashing orange - Y / N`n* lights are cycling between Upstream / Downstream / Online - Y/N`n`nEU is currently receiving a free service. Please complete this order ASAP.`n`nThank you.",

        "HFC - Remote Activation", "Request Type:  Equipment`nSeverity:  Remote Activation Request`n`nHi Team,`n`nPlease proceed with remote activation for [INSERT ORDER ID].`n`nWe can confirm that there is an Arris NTD installed onsite and is powered on. Arris NTD details are as follows:`n[MAC ADDRESS]`n[SERIAL NUMBER]`nArris NTD light status: ON/OFF`n* 4 green light on - Y / N`n* Upstream light is flashing orange - Y / N`n* Downstream light is flashing orange - Y / N`n* lights are cycling between Upstream / Downstream / Online - Y/N`n`nOur customers details are:`n" csTitle "`n" csTitle "x`nXx`n`nPlease remotely activate this service so we can get out EU online.`n`nThank you. ",

        "HFC - Multiple NTDs", "Request Type:  Equipment`nSeverity:  NTD Issue`n`nHi Team,`n`nYour records show that there is multiple Arris NTDs installed onsite at [INSERT LOCID].`n`nIn order for us to place our order against the correct NTD, can you please advise what NTD ID the following Mac & Serial details belong to?`n`n[MAC ADDRESS]`n[SERIAL ADDRESS]`n`nPlease advise so we can proceed with an order for our EU.`n`nThank you",

        "Order Escalation", "Email Subject:`n`nSR INC: [INC ID] | [AUSSIE REF] | Activations | [TECHONOLOGY TYPE i.e. HFC/FTTB]`n`nEmail Body:`n`nOrder Number:  If there is no order ID, please put N/A order`nLOC ID:  [INSERT LOC ID]`nAppointment ID:  If there is no apt ID, please put N/A`nTIO Raised:  If no TIO has been raised please put NOre `nTIO Ref:  If no TIO raised, please put N/A`n`nIssue:`n[Explain what the issue is, generally what is written in the INC]`n`nReason for escalation:`n[Explain why you’re escalating this to the escalations team. Ie. Issue was not resolved/no response within required time frame etc] "
    )

    Global DelaysMap := Map(
        "TC4", "Hi " CustomerName ",`n`nJust reaching out as your order has been delayed. We have received an error from NBN that you are on a higher tier plan with your previous provider. As a result, this would necessitate us closing off that service to open a gigabit plan.`n`nThe main issue is this would likely result in the application needing manual intervention in the morning and as we are not in the office until Monday, will result in significant downtime.`n`nTo alleviate this and get the connection active immediately, there are two options:`n`n1. Please request your current provider immediately downgrade your plan speeds to 100/40 or lower. Once this has been completed, pop into the LiveChat and we can connect your service with us on another port top ensure as little downtime as possible. If choosing this option, please confirm with your provider they have definitely downgraded the plan immediately as we will not be able to order until that has completed.`n`n2. Lower your connected plan with us to get the connection online on another port, close off your old service with your previous provider and then upgrade your plan the day following the closure. This will ensure that your service goes live and that you have minimal downtime also.`n`nEither way, just pop into the LiveChat and we can assist, we are open Mon-Fri 12-8pm AEDT.`n`nIf we haven't heard from you by close of business today, this service won't transfer until Monday evening.",

        "In-Flight Order", "Hey " CustomerName ",`n`nThanks for ordering your NBN with us{!}`n`nWe have had a small hiccup. NBN is telling us there is already an order pending for your premises.`n`nIf you have recently spoken to another retail service provider about an NBN connection, this will probably be what is holding up the process.`n`nIf you wish to go ahead with your broadband order, please:`n`n1. Cancel any orders you may have with other providers.`n2. Let us know by contacting us or responding to this email.`n`nOnce we hear from you, we will re-submit your order immediately!`n`nDue to strict privacy guidelines, we cannot remove pending orders with NBN if they are in place with another service provider, this can only be done with the customer and the service provider in question.`n`nYour order will remain on hold until this order with the other provider has been removed."
    )

    Global HFCMap := Map(
        "Day 1", "",

        "Day 2", "",

        "Day 3", "",

        "Multiple NTDs", "",

        "Signal Issue",""
        )

    Global FTTCMap := Map(
        "Day 1", "",

        "Day 2", "",

        "Day 3", "",
    )

    Global FTTPMap := Map(
        "Multiple UNI-D", "",
        
    )

    Global ValidationMap := Map(
        "Day 1", "",

        "Day 2", "",

        "Day 3", ""
    )

    Global BanlistingMap := Map(
        "Day 1", "",
    )

    Global PaymentMap := Map(
        "Day 1", "",

        "Day 2", "",

        "Day 3", "",
    )

    Global NBNCompMap := Map(
        "Templates", "",
        )

    Global RaisingMap := Map(
        "Templates", "Complaints - Template`n`nIssue: [Describe the issue or problem the customer is facing, providing any relevant background information, incidents, or observations that led to their concerns.]`n`nWhat Has Been Done So Far: [Outline any actions you have taken to address the issue thus far, such as troubleshooting, external escalation, explanation of our policies and procedures, or bringing it to the attention of your supervisor. Have we done everything we can internally to resolve the customer`'s concerns, including following correct escalation pathways, if not, why not?]`n`nCustomer`'s Desired Outcome from the Complaint: [Clearly state what resolution or action that the customer has advised would effectively address the issue.]"
        )

    Global ClarificationMap := Map(
        "Templates", ""
        )

    Global ResolutionsMap := Map(
        "Templates", "",
        )

    Global ChangesMap := Map(
        "Changing PRD", "",
        )

    Global TIOMap := Map(
        "Templates", "",
        )

    Global ContactsMap := Map(
        "SMS", ""
        )

    Global TCSGeneralMap := Map(
        "Availability", "",
        
        "Maintenance", "",

        "Support", "",

        "EU Obligations", "",

        "Data", "",

        "Access", ""
        )

    Global TCSBillingMap := Map(
        "Fees", "",

        "Payment Methods", "",

        "Disputes", "",

        "Other Fees", ""
    )

    Global TCSSuspensionMap := Map(
        "Service Use", "",
        
        "Poor Behaviour", "",

        "Payment Issues", ""
    )

    Global TCSChangesMap := Map(
        "TOS Changes", ""
    )
}
