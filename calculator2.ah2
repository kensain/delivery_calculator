#SingleInstance Force

#Include <tariffs_2023>
#Include <DeliveryCosts>
#Include <AHKv2_Scripts\Info>
#Include <bLib\CBR2>
#Include <bLib\SpellSum>
bCalculator() {

	InitialRates := CBR2()
	date_rate := InitialRates.Date
	EUR_rate := InitialRates.Currency["EUR"]
	CHF_rate := InitialRates.Currency["CHF"]
	USD_rate := InitialRates.Currency["USD"]
	CNY_rate := InitialRates.Currency["CNY"]

	OrderCondition := 0
	SetOrderCondition(Int, *) {
		OrderCondition := Int
		Info(Int)
	}

	g := Gui()
	g.Opt("AlwaysOnTop ToolWindow")
	g.OnEvent("Escape", (*) => close_gui())
	g.OnEvent("Close", (*) => close_gui())
	g.Title := "Калькулятор доставки 2024"
	Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост", "Аммира", "Даты", "Для ДС", "Сумма прописью"])
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
	g.AddText(, "Курсы на")
	CalDate := convert_date(date_rate)
	gDate := g.AddDateTime("yp w108 vCBRDate " . CalDate, "dd.MM.yyyy")
	gDate.OnEvent("Change", (*) => ClickUpdateRates())

	tEUR := g.AddText("r1 w22 xm", "EUR")
	EUR := g.AddEdit("ReadOnly r1 w55 yp", EUR_rate)
	tCHF := g.AddText("r1 w22 yp", "CHF")
	CHF := g.AddEdit("ReadOnly r1 w55 yp", CHF_rate)
	tUSD := g.AddText("r1 w22 xm", "USD")
	USD := g.AddEdit("ReadOnly r1 w55 yp", USD_rate)
	tCNY := g.AddText("r1 w22 yp", "CNY")
	CNY := g.AddEdit("ReadOnly r1 w55 yp", CNY_rate)
	tEUR.OnEvent("DoubleClick", copyText.Bind(EUR))
	tCHF.OnEvent("DoubleClick", copyText.Bind(CHF))
	tUSD.OnEvent("DoubleClick", copyText.Bind(USD))
	tCNY.OnEvent("DoubleClick", copyText.Bind(CNY))
	ProofCheck := g.AddLink("xm", Format('Проверить курс на <a href="https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={1}">сайте</a> ЦБ РФ', date_rate))

	convert_date(date) {
		input := StrSplit(date, ".")
		output := "Choose" . input[3] . input[2] . input[1]
		return output
	}
	UpdateRates(date?) {
		if IsSet(date) = false
			date := "24.08.2022"
		new_rates := CBR2(date)
		date_rate := new_rates.date
		EUR.Text := EUR_rate := new_rates.Currency["EUR"]
		CHF.Text := CHF_rate := new_rates.Currency["CHF"]
		USD.Text := USD_rate := new_rates.Currency["USD"]
		CNY.Text := CNY_rate := new_rates.Currency["CNY"]
		gDate.Value := StrSplit(date_rate, ".")[3] . StrSplit(date_rate, ".")[2] . StrSplit(date_rate, ".")[1]
	}
	ClickUpdateRates(*) {
		new_date := FormatTime(g.Submit(false).CBRDate, "dd.MM.yyyy")
		Sleep(500)
		UpdateRates(new_date)
	}

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
	edit_weight_ammira.OnEvent("Change", (*) => check_state_ammira(edit_weight_ammira))
	; edit_weight_ammira.OnEvent("Change", ObjBindMethod(check_state_ammira, "Call", edit_weight_ammira))
	
	Tab3.UseTab("Даты")
	g.AddRadio("visOffer Checked1", "КП клиенту").OnEvent("Click", switchRadio.Bind("toOffer"))
	g.AddRadio("visOrder Checked0", "Размещение заказа").OnEvent("Click", switchRadio.Bind("toOrder"))
	SourceDate := g.AddText("r1", "Дата КП:")
	g.AddDateTime("yp-3 x75 vStartDate w97", "dd.MM.yyyy")
	TextDays := g.AddText("x22 y114 r1 w150", "+/- дней(EXW):")
	g.AddEdit("w150")
	days_edit := g.AddUpDown("vDays Range0-180", 1)
	TextWeeks := g.AddText("r1 w150", "+/- недель(DDP):")
	g.AddEdit("w150")
	weeks_edit := g.AddUpDown("vWeeks Range0-180", 11)
	days_edit.OnEvent("Change", (*) => check_state_date())
	; days_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
	weeks_edit.OnEvent("Change", (*) => check_state_date())
	; weeks_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
	g.AddRadio("vConditionGroup Checked", "От аванса").OnEvent("Click", SetOrderCondition.Bind(1))
	g.AddRadio("", "От подписания").OnEvent("Click", SetOrderCondition.Bind(2))
	g.AddRadio("", "От размещения заказа").OnEvent("Click", SetOrderCondition.Bind(3))
	date_button := g.AddButton("w150 Default", "Рассчитать")
	date_button.OnEvent("Click", (*) => when_clicked())
	g.AddGroupBox("w150 h60", "Сроки поставки")
	resultText1A := g.AddText("xp+5 yp+17 w70")
	resultText1A.OnEvent("DoubleClick", copyText.Bind(resultText1A))
	resultText1B := g.AddText("x90 yp w70")
	resultText1B.OnEvent("DoubleClick", copyText.Bind(resultText1B))
	resultText2A := g.AddText("x27 yp+20 w70")
	resultText2A.OnEvent("DoubleClick", copyText.Bind(resultText2A))
	resultText2B := g.AddText("x90 yp w70")
	resultText2B.OnEvent("DoubleClick", copyText.Bind(resultText2B))
	date := {
		StartDate: "",
		Weeks: "",
		Days: ""
	}
	Tab3.UseTab("Для ДС")
	g.AddText(, "Точка отсчёта")
	g.AddDateTime("vAddStart w150", "dd.MM.yyyy")
	g.AddText(, "Изначальный срок (нед.)")
	g.AddEdit("w150")
	g.AddUpDown("vOrigTime Range0-180", 1)
	g.AddText(, "Новая дата готовности")
	g.AddDateTime("vNewBuzDate w150", "dd.MM.yyyy")
	g.AddText(, "Недель от BUZ до клиента:")
	g.AddEdit("w150")
	g.AddUpDown("vFcaDdp Range0-180", 11)
	AddButton := g.AddButton("w150 Default", "Рассчитать")
	AddButton.OnEvent("Click", (*) => addClick())
	g.AddGroupBox("w150 h130", "Новый срок поставки")
	; newDelivery := g.AddText("xp+5 yp+16 w125", "Новый срок поставки")
	newDate := g.AddText("xp+5 yp+20 w125")
	newDate.OnEvent("DoubleClick", copyText.Bind(newDate))
	newDateWeeks := g.AddText("w125")
	newDateWeeks.OnEvent("DoubleClick", copyText.Bind(newDateWeeks))

	addClick() {
		startDate := g.Submit(0).AddStart
		origWeeks := g.Submit(0).OrigTime
		NewBuzDate := g.Submit(0).NewBuzDate
		FcaDdpWeeks := g.Submit(0).FcaDdp

		; Первоначальный срок поставки (long date)
		originalDeliveryDate := DateAdd(startDate, origWeeks*7, "Days")
		; Первоначальный срок поставки (формат)
		originalDeliveryDateF := FormatTime(originalDeliveryDate, "dd.MM.yyyy")
		; Первоначальный срок поставки в неделях
		originalDeliveryTimeWeeks := Integer(DateDiff(originalDeliveryDate, startDate, "Days") / 7)
		; Новый срок поставки (long date)
		newDDPDate := DateAdd(NewBuzDate, FcaDdpWeeks*7, "Days")
		; Новый срок поставки (формат)
		newDDPDateF := FormatTime(newDDPDate, "dd.MM.yyyy")
		newDDPDateWeeks := Ceil(DateDiff(newDDPDate, startDate, "Days") / 7)
		; MsgBox("Первоначальный срок поставки: " originalDeliveryDateF " (" originalDeliveryTimeWeeks " нед.)`nНовый срок поставки: " newDDPDateF " (" newDDPDateWeeks " нед.)")

		ControlSetText(newDDPDateF, newDate)
		ControlSetText(newDDPDateWeeks " недель", newDateWeeks)
	}

	Tab3.UseTab("Сумма прописью")
	g.AddText(, "Сумма:")
	g.AddEdit("r1 w150 vSum")
	g.AddRadio("Checked vCurrencyRadioGroup", "RUB")
	g.AddRadio("yp", "EUR")
	g.AddRadio("yp", "USD")
	g.AddText("x22 y120 w100", "НДС:")
	vat := [0, 7, 10, 12, 13, 15, 17, 18, 20]
	g.AddDropDownList("x55 y116 r9 Choose9 w117 vTax", vat)
	spell_button := g.AddButton("x22 y145 w150", "Превратить в текст")
	spell_button.OnEvent("Click", (*) => ClickSpell())
	sum_field := g.AddEdit("w150 r10 ReadOnly")

	ClickSpell() {
		SpellData := {
			digit_sum: RegExReplace(g.Submit(0).Sum, "[A-Za-z\s]*"),
			currency: g.Submit(0).CurrencyRadioGroup,
			vat: g.Submit(0).Tax
		}
		if SpellData.digit_sum = "" {
			ControlSetText("", sum_field)
			return
		}
		switch SpellData.currency{
			case 1: currency := "RUB"
			case 2: currency := "EUR"
			case 3: currency := "USD"
		}
		spelt_sum := SpellSum(SpellData.digit_sum, currency, SpellData.vat)
		ControlSetText(spelt_sum, sum_field)
	}

	Tab3.UseTab()

	/**
	 * 
	 * @param {String} Mode 'Date' or 'Rate'
	 * @param {String} textToCopy Text to copy
	 */
	copyText(Control, *) {
		if ControlGetText(ControlGetHwnd(Control)) = ""
			return
		Conditions := Map(
			1, "{1} с аванса",
			2, "{1} от подписания",
			3, "{1} от размещения",
		)
		if OrderCondition != 0
			A_Clipboard := Format(Conditions[OrderCondition], ControlGetText(ControlGetHwnd(Control)))
		else
			A_Clipboard := ControlGetText(ControlGetHwnd(Control))
		textToFormat := "Текст '{1}' скопирован!"
		Info(Format(textToFormat, A_Clipboard))
	}
	
	button_ammira_event(*) {
		up_to_500kg := 1
		up_to_1t := 2
		up_to_2t := 3
		up_to_3t := 4
		up_to_5t := 5
		up_to_10t := 6
		up_to_15t := 7
		up_to_20t := 8
		
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
			MsgBox("Вес превышает лимит в 20 тонн!", "Ошибка!", "0x30")
			Exit()
		}
		
		price := ammira[destination][tariff]

		costs := [price, convertPrice(price, EUR_rate), convertPrice(price, CHF_rate), convertPrice(price, USD_rate), convertPrice(price, CNY_rate)]
		title := Format("Стоимость доставки {1} кг в {2}", weight, destination)
		DeliveryCosts(title, costs*)
		; MsgBox(Format("{1} рублей`n`nПри пересчёте в валюту:`n{2} евро`n{3} франков`n{4} долларов`n{5} юаней", price, convertPrice(price, EUR_rate), convertPrice(price, CHF_rate), convertPrice(price, USD_rate), convertPrice(price, CNY_rate)), "Стоимость доставки " weight " кг в " destination, "0x40")
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

	switchRadio(calculationMode, *) {
		switch calculationMode {
			case "toOffer":
				Offer()
			case "toOrder":
				Order()
		}
		
		Offer() {
			ControlSetText("Дата КП", SourceDate)
			ControlSetText("+/- дней(EXW):", TextDays)
			ControlSetText("+/- недель(DDP):", TextWeeks)
			days_edit.Value := 1
			weeks_edit.Value := 11
			SetOrderCondition(g.Submit(0).ConditionGroup)
		}
		Order() {
			ControlSetText("Дата заказа", SourceDate)
			ControlSetText("Недель от BUZ до клиента:", TextDays)
			ControlSetText("Общий срок поставки:", TextWeeks)
			days_edit.Value := 11
			weeks_edit.Value := 12
			SetOrderCondition(0)
		}

	}

	when_clicked() {
		isOffer := g.Submit(0).isOffer
		isOrder := g.Submit(0).isOrder

		SetOrderCondition(g.Submit(0).ConditionGroup)

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

			if isOffer = true
				calculateOffer()
			else if isOrder = true
				calculateOrder()

			calculateOffer() {
				deliveryTime := (Ceil(number_of_days / 5) * 7) + number_of_weeks
				calculatedDate := DateAdd(date.StartDate, deliveryTime, "Days"), "dd.MM.yyyy" ;25.09.2023
				deliveryDate := FormatTime(DateAdd(date.StartDate, deliveryTime, "Days"), "dd.MM.yyyy") ;25.09.2023
				deliveryWeeks := Ceil(DateDiff(calculatedDate, date.StartDate, "Days") / 7)

				ControlSetText(deliveryDate, resultText1A)
				ControlSetText(deliveryWeeks " недель", resultText1B)
				ControlSetText("", resultText2A)
				ControlSetText("", resultText2B)
			}
			
			calculateOrder() {
				weeksToRussia := number_of_days * 7
				totalWeeks := number_of_weeks
				calculationDDP := DateAdd(date.StartDate, totalWeeks, "Days")
				calculationFCA := DateAdd(calculationDDP, -weeksToRussia, "Days")
				calculationEXW := DateAdd(calculationDDP, -calculationFCA, "Days")
				FCA := FormatTime(calculationFCA, "dd.MM.yyyy")
				DDP := FormatTime(calculationDDP, "dd.MM.yyyy")
				ControlSetText("Дата EXW:", resultText1A)
				ControlSetText(FCA, resultText1B)
				ControlSetText("Дата DDP:", resultText2A)
				ControlSetText(DDP, resultText2B)
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
				price := tariffs[destination][tariff] + ((weight - countStart) * tariffs[destination][markup])
		}
		
		costs := [price, convertPrice(price, EUR_rate), convertPrice(price, CHF_rate), convertPrice(price, USD_rate), convertPrice(price, CNY_rate)]
		title := Format("Стоимость доставки {1} кг в {2}", weight, tariffs[destination][1])
		if price != "****"
			DeliveryCosts(title, costs*)
		else {
			Result := MsgBox(Format("К сожалению, для отправлений в {1} свыше 32 кг действует специальный тариф с применением регрессивной шкалы за каждый следующий кг, поэтому надо пересчитывать на сайте. `nОткрыть калькулятор на сайте?", tariffs[destination][1]), Format("Отправка в {}", tariffs[destination][1]), "YesNo")
			if Result = "Yes"
				Run("https://garantpost.ru/tools")
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
}
bCalculator()