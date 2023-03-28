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
g.OnEvent("Escape", (*) => close_gui())
g.OnEvent("Close", (*) => close_gui())
g.AddText(, "Калькулятор доставки")
Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост","Аммира", "Даты"])
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
editWeight.OnEvent("Change", ObjBindMethod(check_state_garantpost, "Call", editWeight))
Tab3.UseTab()
g.AddGroupBox("h60 w175", "Курсы на " date_rate)
g.AddText("right xp+5 yp+17", "EUR: " EUR_rate "`nCHF: " CHF_rate "`nUSD: " USD_rate)

Tab3.UseTab("Аммира")
ammiraList := []
if ammira.Count = 0
    ammiraList.Push("Coming soon!")
else
    for key, value in ammira
        ammiraList.Push(key)
g.AddListBox("r20 vAmmiraChoice Choose1 w150", ammiraList)
g.AddText(, "Вес в КГ:")
edit_weight_ammira := g.AddEdit("r1 vWeightAmmira w150 Number")
button_ammira := g.AddButton("Default w150 Disabled", "OK")
button_ammira.OnEvent("Click", button_ammira_event)
edit_weight_ammira.OnEvent("Change", ObjBindMethod(check_state_ammira, "Call", edit_weight_ammira))

Tab3.UseTab("Даты")
g.AddText(, "Start date")
g.AddDateTime("vStartDate w150", "dd.MM.yyyy")
g.AddText(, "+/- Days:")
days_edit := g.AddEdit("r1 w150 vDays Number", 1)
g.AddText(, "+/- Weeks:")
weeks_edit := g.AddEdit("r1 w150 vWeeks Number", 11)
days_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
weeks_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
isDelivery := g.AddCheckbox("vDeliveryChecked Checked", "Calculate EXW + DDP")
date_button := g.AddButton("w150 Default", "Calculate")
date_button.OnEvent("Click", (*) => when_clicked())
g.AddGroupBox("w150 h150", "Result:")
date_result := g.AddText("xp+5 yp+16 w125", "The result will be printed here")
date := {
    StartDate: "",
    Weeks: "",
    Days: ""
}
; delivery_calculation_checked := ""
Tab3.UseTab()

button_ammira_event(*) {
    up_to_500kg := 1
    up_to_1t    := 2
    up_to_2t    := 3
    up_to_3t    := 4
    up_to_5t    := 5
    up_to_10t   := 6
    up_to_15t   := 7
    up_to_20t   := 8

    destination := g.Submit().AmmiraChoice
    weight := g.Submit().WeightAmmira

    if weight <= 500
        tariff := up_to_500kg
    else if weight <= 1000
        tariff := up_to_1t
    else if weight <= 2000
        tariff := up_to_2t
    else if weight <= 3000
        tariff := up_to_3t
    else if weight <= 5000
        tariff := up_to_5t
    else if weight <= 10000
        tariff := up_to_10t
    else if weight <= 15000
        tariff := up_to_15t
    else if weight <= 20000
        tariff := up_to_20t
    else {
        MsgBox("Вес превышает лимит в 20 тонн!")
        Exit()
    }

    price := ammira[destination][tariff]

    MsgBox(price " рублей`n`nПри пересчёте в валюту:`n" convertPrice(price, EUR_rate) " евро`n" convertPrice(price, CHF_rate) " франков`n" convertPrice(price, USD_rate) " долларов", "Стоимость доставки " weight " кг в " destination)
}

check_state_garantpost(edit, *) {
    if edit.Value != ""
        button.Enabled := 1
    else
        button.Enabled := 0
}
check_state_ammira(edit, *) {
    if edit.Value != ""
        button_ammira.Enabled := 1
    else
        button_ammira.Enabled := 0
}
check_state_date(*) {
    if (weeks_edit.Value != "" or days_edit.Value != "")
        date_button.Enabled := 1
    else
        date_button.Enabled := 0
}

when_clicked() {

    delivery_calculation_checked := g.Submit(0).DeliveryChecked

    date.StartDate := g.Submit(0).StartDate
    date.Weeks := g.Submit(0).Weeks
    date.Days := g.Submit(0).Days
    calculate()

    calculate() {
        
        if date.Weeks != ""
            number_of_weeks := Integer(date.Weeks) * 7
        else
            number_of_weeks := 0
        if date.Days != ""
            number_of_days := Integer(date.Days)
        else
            number_of_days := 0
    
        switch delivery_calculation_checked {
            case 1:
                calculate_delivery()
            case 0:
                calculate_date()
            default:
                MsgBox("The value of delivery_calculation_checked is something different.")
        }
            if delivery_calculation_checked = 1
                calculate_delivery()
            else
                calculate_date()
    
        calculate_delivery() {
            exw := DateAdd(date.StartDate, number_of_days, "Days")
            date_exw := FormatTime(exw, "dd.MM.yyyy")
            ddp := DateAdd(exw, number_of_weeks, "Days")
            date_ddp := FormatTime(ddp, "dd.MM.yyyy")
    
            result := "EXW: " date_exw "`nDDP: " date_ddp
    
            ControlSetText(result, date_result)
        }
    
        calculate_date() {
            new_date := DateAdd(date.StartDate, number_of_days + number_of_weeks, "Days")
            fDate_exw := FormatTime(new_date, "dd.MM.yyyy")
    
            result := "New date: " fDate_exw
    
            ControlSetText(result, date_result)
        }
    
    }
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

    if price != "****"
        MsgBox(price " рублей`n`nПри пересчёте в валюту:`n" convertPrice(price, EUR_rate) " евро`n" convertPrice(price, CHF_rate) " франков`n" convertPrice(price, USD_rate) " долларов", "Стоимость доставки " weight " кг в " tariffs[destination][1])
    else {
        Result := MsgBox("К сожалению, для отправлений в " tariffs[destination][1] " свыше 32 кг действует специальный тариф с применением регрессивной шкалы за каждый следующий кг, поэтому надо пересчитывать на сайте. Открыть калькулятор на сайте?", "Отправка в " tariffs[destination][1], "YesNo")
        if Result = "Yes"
            Run "https://garantpost.ru/tools"
        else
            return
    }
}

close_gui() {
    g.Hide(`)
    Exit()
}

toDot(number_with_comma) {
    return StrReplace(number_with_comma, ",", ".")
}

toComma(number_with_comma) {
    return StrReplace(number_with_comma, ".", ",")
}

convertPrice(price, currency) {
    return toComma(Format("{:.2f}", price / toDot(currency)))
}
g.Show()

; Win+Shift+D to show the GUI
#+d::g.Show()
; MenuHandler(Item, *) {
;     MsgBox "You selected " Item
; }