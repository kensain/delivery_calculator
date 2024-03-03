#Include <Info>
Class DeliveryCosts {
    __New(Title, CostRUB, CostEUR, CostCHF, CostUSD, CostCNY) {
        gCosts := Gui()
        gCosts.Title := Title
        gCosts.Opt("-MinimizeBox -MaximizeBox +AlwaysOnTop +ToolWindow")
        gCosts.OnEvent("Escape", __OnEventHandler)
        gCosts.OnEvent('Close', __OnEventHandler)

        gCosts.Add("Text", "vCalculatedCostRUB x22 y16 w120 h23 +0x200", Format("{} рублей", CostRUB))
        gCosts.Add("Text", "vCalculatedCostRUB_rounded x22 y35 w120 h23 +0x200", Format("{} рублей (округл.)",  Round(CostRUB + 100, -2)))
        gCosts.Add("Text", "x22 y56 w129 h23 +0x200", "При пересчёте в валюту:")
        gCosts.Add("Text", "vCalculatedCostEUR x22 y80 w120 h23 +0x200", Format("{} евро", CostEUR))
        gCosts.Add("Text", "vCalculatedCostCHF x22 y104 w120 h23 +0x200", Format("{} франков", CostCHF))
        gCosts.Add("Text", "vCalculatedCostUSD x22 y128 w120 h23 +0x200", Format("{} долларов", CostUSD))
        gCosts.Add("Text", "vCalculatedCostCNY x22 y152 w120 h21 +0x200", Format("{} юаней", CostCNY))
        
        gCosts["CalculatedCostRUB"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostRUB"))
        gCosts["CalculatedCostRUB_rounded"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostRUB_rounded"))
        gCosts["CalculatedCostEUR"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostEUR"))
        gCosts["CalculatedCostCHF"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostCHF"))
        gCosts["CalculatedCostUSD"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostUSD"))
        gCosts["CalculatedCostCNY"].OnEvent("DoubleClick", (*) => __CopyText("CalculatedCostCNY"))

        gCosts.AddButton("vButton x22 y184 w80 h23 +Default", "&OK")
        gCosts["Button"].OnEvent("Click", __OnEventHandler)
        
        gCosts.Show("w300 h215")
        
        __OnEventHandler(*) {
            gCosts.Destroy()
        }
        
        __CopyText(vControlName) {
            A_Clipboard := RegExReplace(gCosts[vControlName].Text, " .*")
            Info(Format("Стоимость '{}' скопирована!", A_Clipboard))
            gCosts.Destroy()
        }
    }
}