#SingleInstance Force

#Include tariffs_2023.ah2
#Include <CBR>

date_rate := CBR("EUR").date
EUR_rate := CBR("EUR").rate
CHF_rate := CBR("CHF").rate
USD_rate := CBR("USD").rate
CNY_rate := CBR("CNY").rate

g := Gui()
g.Opt("AlwaysOnTop ToolWindow")
g.OnEvent("Escape", (*) => close_gui())
g.OnEvent("Close", (*) => close_gui())
g.Title := "Калькулятор доставки 2023"
Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост","Аммира", "Даты"])
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
; SelectedDate := g.AddGroupBox("h60 w175", "Курсы на " date_rate)
g.AddText(, "Курсы на")
CalDate := convert_date(date_rate)
gDate := g.AddDateTime("yp w108 vCBRDate " . CalDate, "dd.MM.yyyy")
; Rates := g.AddText("right xp+5 yp+17", "EUR: " EUR_rate "`nCHF: " CHF_rate "`nUSD: " USD_rate)
; Rates := g.AddText("xm", "EUR: " EUR_rate "`nCHF: " CHF_rate "`nUSD: " USD_rate)
; g.AddText("xm", "EUR")
; EUR := g.AddEdit("ReadOnly")
; g.AddText(, "CHF")
; CHF := g.AddEdit("ReadOnly")
; g.AddText(, "USD")
; USD := g.AddEdit("ReadOnly")
; g.AddText(, "CNY")
; CNY := g.AddEdit("ReadOnly")
gDate.OnEvent("Change", (*) => update_rateButton())

g.AddEdit("ReadOnly Disabled r1 w28 xm", "EUR")
EUR := g.AddEdit("ReadOnly r1 w46 yp", EUR_rate)
g.AddEdit("ReadOnly Disabled r1 w28 yp", "CHF")
CHF := g.AddEdit("ReadOnly r1 w46 yp", CHF_rate)
g.AddEdit("ReadOnly Disabled r1 w28 xm", "USD")
USD := g.AddEdit("ReadOnly r1 w46 yp", USD_rate)
g.AddEdit("ReadOnly Disabled r1 w28 yp", "CNY")
CNY := g.AddEdit("ReadOnly r1 w46 yp", CNY_rate)
ProofCheck := g.AddLink("xm", 'Проверить курс на <a href="https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To=' . date_rate . '">сайте</a> ЦБ РФ')

convert_date(date) {
    input := StrSplit(date, ".")
    output := "Choose" . input[3] . input[2] . input[1]
    return output
}
update_rate(date?) {
    if IsSet(date) = false
        date := "24.08.2022"
    
    date_rate := CBR("EUR", , date).date
    EUR.Text := EUR_rate := CBR("EUR",  ,date).rate
    CHF.Text := CHF_rate := CBR("CHF",  ,date).rate
    USD.Text := USD_rate := CBR("USD",  ,date).rate
    CNY.Text := CNY_rate := CBR("CNY",  ,date).rate
    gDate.Value := StrSplit(date_rate, ".")[3] . StrSplit(date_rate, ".")[2] . StrSplit(date_rate, ".")[1]
}
update_rateButton(*) {
    d := FormatTime(g.Submit(false).CBRDate, "dd.MM.yyyy")
    Sleep(500)
    update_rate(d)
}
SB := g.AddStatusBar(, "...")
; SB.SetText("There are " . EUR_rate . " rows selected.")

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
g.AddText(, "Исходная дата")
g.AddDateTime("vStartDate w150", "dd.MM.yyyy")
g.AddText(, "+/- дней(EXW):")
days_edit := g.AddEdit("r1 w150 vDays Number", 1)
g.AddText(, "+/- недель(DDP):")
weeks_edit := g.AddEdit("r1 w150 vWeeks Number", 11)
days_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
weeks_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
isDelivery := g.AddCheckbox("vDeliveryChecked Checked", "Рассчитать EXW + DDP")
date_button := g.AddButton("w150 Default", "Рассчитать")
date_button.OnEvent("Click", (*) => when_clicked())
g.AddGroupBox("w150 h150", "Результат:")
date_result := g.AddText("xp+5 yp+16 w125 r10", "Результат будет здесь")
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

    MsgBox(price " рублей`n`nПри пересчёте в валюту:`n" convertPrice(price, EUR_rate) " евро`n" convertPrice(price, CHF_rate) " франков`n" convertPrice(price, USD_rate) " долларов`n" convertPrice(price, CNY_rate) " юаней", "Стоимость доставки " weight " кг в " destination)
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
            case true:
                calculate_delivery()
            case false:
                calculate_date()
            default:
                MsgBox("The value of delivery_calculation_checked is something different.")
        }

        if delivery_calculation_checked = true
            calculate_delivery()
        else
            calculate_date()
    
        calculate_delivery() {

            if number_of_days < 5
                exw := DateAdd(date.StartDate, number_of_days, "Days")
            else
                exw := DateAdd(date.StartDate, ((number_of_days // 5) * 7), "Days")
            date_exw := FormatTime(exw, "dd.MM.yyyy")
            ddp := DateAdd(exw, number_of_weeks, "Days")
            date_ddp := FormatTime(ddp, "dd.MM.yyyy")
            total_weeks := DateDiff(ddp, date.StartDate, "Days") // 7
    
            result := "EXW: " date_exw "`nDDP: " date_ddp "`nИтого: " total_weeks " нед."
    
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
    g.Destroy()
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