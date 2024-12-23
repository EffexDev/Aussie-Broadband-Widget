; The series of functions below are what generates the templates when the generate button is clicked. 
; They are all the same so I will break down the logic of the first one and it can be applied to the rest.
RunGeneral(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedGeneralReason)
        Output := templateMap.Get(Saved.PickedGeneral)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}

RunAccount(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedAccountReason)
        Output := templateMap.Get(Saved.PickedAccount)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}

RunFault(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedFaultReason)
        Output := templateMap.Get(Saved.PickedFault)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}

RunDelivery(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedDeliveryReason)
        Output := templateMap.Get(Saved.PickedDelivery)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}

RunComplaint(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedComplaintReason)
        Output := templateMap.Get(Saved.PickedComplaint)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}

RunTCS(ctrl, *) {
    Global ShowNotesButton
    Saved := AussieGui.Submit(False)
    Output := ""

    UpdateTemplates()

    try {
        templateMap := TemplatesMaps.Get(Saved.PickedTCSReason)
        Output := templateMap.Get(Saved.PickedTCS)

        showNotes := AussieGui["ShowNotesButton"].Value
        
        if (showNotes) {
            ToolsTab.Choose(1)
            ControlFocus NotePadEmbedded
            NotePadEmbedded.Focus()
            Send Output
        } else {
            TemplatesPad()
            ToolsTab.Choose(1)
            ControlFocus Templates
            Templates.Focus()
            Send Output
        }
    } catch as Error {
        MsgBox "Make sure you select all options and enter customer name."
    }
}