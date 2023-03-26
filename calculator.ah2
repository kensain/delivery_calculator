#SingleInstance Force

#Include tariffs_2023.ah2

; Create WinHttpRequest object
WinHttp := ComObject("WinHttp.WinHttpRequest.5.1")

; Download XML data from the URL
url := "http://www.cbr.ru/scripts/XML_daily.asp"
WinHttp.Open("GET", url)
WinHttp.Send()

if (WinHttp.Status != 200) {
    MsgBox("Error", "Failed to download XML data.`n`nMake sure you have an active internet connection and try again.")
    ExitApp
}

xml := WinHttp.ResponseText

; Parse XML data
XmlDoc := ComObject("Msxml2.DOMDocument.6.0")
XmlDoc.async := false
XmlDoc.loadXML(xml)

if (XmlDoc.parseError.errorCode != 0) {
    MsgBox("Error", "Failed to parse XML data.`n`nError: " . XmlDoc.parseError.reason)
    ExitApp
}

; Extract currency data
root := XmlDoc.documentElement
currencyNodes := root.getElementsByTagName("Valute")
date_rate := root.getAttribute("Date")

; Initialize variables for storing currency rates
EUR_rate := ""
CHF_rate := ""
USD_rate := ""

for each in currencyNodes {
    charCode := each.getElementsByTagName("CharCode").item(0).text
    if (charCode = "EUR") {
        EUR_rate := each.getElementsByTagName("Value").item(0).text
    } else if (charCode = "CHF") {
        CHF_rate := each.getElementsByTagName("Value").item(0).text
    } else if (charCode = "USD") {
        USD_rate := each.getElementsByTagName("Value").item(0).text
    }
}

g := Gui()
g.Opt("-MaximizeBox AlwaysOnTop ToolWindow")
g.AddText(, "Калькулятор доставки")
Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост","Аммира"])
; g.Add("Tab3", "vTab3", ["Tab no.1","Tab no.2"])
; g.AddDropDownList("vGarantpostChoice")
garantpostList := []
for each in tariffs
    garantpostList.Push(each[1])
g.AddListBox("r20 vGarantpostChoice Choose1 w150 AltSubmit", garantpostList)
g.AddText(, "Вес в КГ:")
editWeight := g.AddEdit("r1 vWeight w150 Number")
button := g.AddButton("Default w150 Disabled", "OK")
button.OnEvent("Click", buttonEvent)
editWeight.OnEvent("Change", ObjBindMethod(checkEditState, "Call", editWeight))
Tab3.UseTab()
g.AddGroupBox("h60 w175", "Курсы на " date_rate)
g.AddText("right xp+5 yp+17", "EUR: " EUR_rate "`nCHF: " CHF_rate "`nUSD: " USD_rate)

Tab3.UseTab("Аммира")
ammiraList := []
if ammira.Count = 0
    ammiraList.Push("Coming soon!")
g.AddListBox("r20 vAmmiraChoice Choose1 w150 AltSubmit", ammiraList)
g.AddText(, "Вес в КГ:")
Tab3.UseTab()

checkEditState(edit, *) {
    if edit.Value != ""
        button.Enabled := 1
    else
        button.Enabled := 0
}

buttonEvent(*) {
    destination := g.Submit().GarantpostChoice
    weight := g.Submit().Weight

    ; Conditions:
    if weight <= 0.1
        tariff := 2
    else if weight <= 0.5
        tariff := 3
    else if weight <= 1
        tariff := 4
    else if weight > 1 and weight < 32 {
        tariff := 4
        markup := 5
        countStart := 1
    }
    else if weight = 32
        tariff := 6
    else if weight > 32 {
        tariff := 6
        markup := 7
        countStart := 32
    }
    else
        MsgBox("Some error - the weight is " weight)

    ; Calculate price
    if !IsSet(markup)
        price := tariffs[destination][tariff]
    else {
        if tariffs[destination][markup] = "****"
            price := "****"
        else
            price := tariffs[destination][tariff] + ((weight - countStart)*tariffs[destination][markup])      
    }

    toDot(number_with_comma) {
        return StrReplace(number_with_comma, ",", ".")
    }

    toComma(number_with_comma) {
        return StrReplace(number_with_comma, ".", ",")
    }

    convertPrice(currency) {
        return toComma(Format("{:.2f}", price / toDot(currency)))
    }

    if price != "****"
        MsgBox(price " рублей`n`nПри пересчёте в валюту:`n" convertPrice(EUR_rate) " евро`n" convertPrice(CHF_rate) " франков`n" convertPrice(USD_rate) " долларов", "Стоимость доставки " weight " кг в " tariffs[destination][1])
    else {
        Result := MsgBox("К сожалению, для отправлений в " tariffs[destination][1] " свыше 32 кг действует специальный тариф с применением регрессивной шкалы за каждый следующий кг, поэтому надо пересчитывать на сайте. Открыть калькулятор на сайте?", "Отправка в " tariffs[destination][1], "YesNo")
        if Result = "Yes"
            Run "https://garantpost.ru/tools"
        else
            return
    }
}
g.Show()

; Win+Shift+D to show the GUI
#+d::g.Show()
; MenuHandler(Item, *) {
;     MsgBox "You selected " Item
; }