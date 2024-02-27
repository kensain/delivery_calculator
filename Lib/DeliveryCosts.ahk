#Include <Info>
Class DeliveryCosts {

    __New(Title, CostRUB, CostEUR, CostCHF, CostUSD, CostCNY) {
        gCosts := Gui()
        gCosts.Title := Title
        gCosts.Opt("-MinimizeBox -MaximizeBox +AlwaysOnTop +ToolWindow")
        TextCalculatedCostRUB := gCosts.Add("Text", "vCalculatedCostRUB x22 y16 w120 h23 +0x200", Format("{} рублей", CostRUB))
        TextCalculatedCostRUB_rounded := gCosts.Add("Text", "vCalculatedCostRUB_rounded x22 y35 w120 h23 +0x200", Format("{} рублей (округл.)",  Round(CostRUB + 100, -2)))
        gCosts.Add("Text", "x22 y56 w129 h23 +0x200", "При пересчёте в валюту:")
        TextCalculatedCostEUR := gCosts.Add("Text", "vCalculatedCostEUR x22 y80 w120 h23 +0x200", Format("{} евро", CostEUR))
        TextCalculatedCostCHF := gCosts.Add("Text", "vCalculatedCostCHF x22 y104 w120 h23 +0x200", Format("{} франков", CostCHF))
        TextCalculatedCostUSD := gCosts.Add("Text", "vCalculatedCostUSD x22 y128 w120 h23 +0x200", Format("{} долларов", CostUSD))
        TextCalculatedCostCNY := gCosts.Add("Text", "vCalculatedCostCNY x22 y152 w120 h21 +0x200", Format("{} юаней", CostCNY))
        
        TextCalculatedCostRUB.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostRUB)))
        TextCalculatedCostRUB_rounded.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostRUB_rounded)))
        TextCalculatedCostEUR.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostEUR)))
        TextCalculatedCostCHF.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostCHF)))
        TextCalculatedCostUSD.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostUSD)))
        TextCalculatedCostCNY.OnEvent("DoubleClick", (*) => CopyText(ControlGetText(TextCalculatedCostCNY)))

        ButtonBtnOk := gCosts.Add("Button", "vBtnOk x22 y184 w80 h23 +Default", "&OK")
        ButtonBtnOk.OnEvent("Click", OnEventHandler)
        gCosts.OnEvent('Close', (*) => ExitApp())
        gCosts.Show("w300 h215")
        
        OnEventHandler(*) {
            gCosts.Destroy()
        }
        
        CopyText(Text) {
            A_Clipboard := RegExReplace(Text, " .*")
            Info(Format("Стоимость '{}' скопирована!", A_Clipboard))
        }
    }
}