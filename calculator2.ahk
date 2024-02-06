#SingleInstance Force
#Requires AutoHotkey v2.0-a
#Include <tariffs_2023>
#Include <DeliveryCosts>
#Include <AHKv2_Scripts\Info>
#Include <bLib\CBR2>
#Include <bLib\SpellSum>
bCalculator() {

	initial_rates := CBR2()
	date_rate := initial_rates.Date
	EUR_rate := initial_rates.Currency["EUR"]
	CHF_rate := initial_rates.Currency["CHF"]
	USD_rate := initial_rates.Currency["USD"]
	CNY_rate := initial_rates.Currency["CNY"]

	order_condition := 1
	set_order_condition(Int, *) {
		order_condition := Int
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
	edit_weight := g.AddEdit("r1 vWeight w150 Number")
	button := g.AddButton("Default w150 Disabled", "OK")
	button.OnEvent("Click", button_event)
	edit_weight.OnEvent("Change", ObjBindMethod(check_state_garantpost, "Call", edit_weight))
	Tab3.UseTab()
	g.AddText(, "Курсы на")
	calendar_date := convert_date(date_rate)
	gDate := g.AddDateTime("yp w108 vCBRDate " . calendar_date, "dd.MM.yyyy")
	gDate.OnEvent("Change", (*) => click_update_rates())

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
	proof_check := g.AddLink("xm", Format('Проверить курс на <a href="https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={1}">сайте</a> ЦБ РФ', date_rate))

	convert_date(date) {
		input := StrSplit(date, ".")
		output := "Choose" . input[3] . input[2] . input[1]
		return output
	}
	update_rates(date?) {
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
	click_update_rates(*) {
		new_date := FormatTime(g.Submit(false).CBRDate, "dd.MM.yyyy")
		Sleep(500)
		update_rates(new_date)
	}

	Tab3.UseTab("Аммира")
	ammira_list := []
	if ammira.Count = 0
		ammira_list.Push("Coming soon!")
	else
		for key, value in ammira
			ammira_list.Push(key)
	g.AddListBox("r20 vAmmiraChoice Choose1 w150", ammira_list)
	g.AddText(, "Вес в КГ:")
	edit_weight_ammira := g.AddEdit("r1 vWeightAmmira w150 Number")
	button_ammira := g.AddButton("Default w150 Disabled", "OK")
	button_ammira.OnEvent("Click", button_ammira_event)
	edit_weight_ammira.OnEvent("Change", (*) => check_state_ammira(edit_weight_ammira))
	; edit_weight_ammira.OnEvent("Change", ObjBindMethod(check_state_ammira, "Call", edit_weight_ammira))
	
	Tab3.UseTab("Даты")
	g.AddRadio("vis_offer Checked1", "КП клиенту").OnEvent("Click", switch_radio.Bind("toOffer"))
	g.AddRadio("vis_order Checked0", "Размещение заказа").OnEvent("Click", switch_radio.Bind("toOrder"))
	source_date := g.AddText("r1", "Дата КП:")
	g.AddDateTime("yp-3 x75 vStart_Date w97", "dd.MM.yyyy")
	text_days := g.AddText("x22 y114 r1 w150", "+/- дней(EXW):")
	g.AddEdit("w150")
	days_edit := g.AddUpDown("vDays Range0-180", 1)
	text_weeks := g.AddText("r1 w150", "+/- недель(DDP):")
	g.AddEdit("w150")
	weeks_edit := g.AddUpDown("vWeeks Range0-180", 11)
	days_edit.OnEvent("Change", (*) => check_state_date())
	; days_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
	weeks_edit.OnEvent("Change", (*) => check_state_date())
	; weeks_edit.OnEvent("Change", ObjBindMethod(check_state_date, "Call"))
	r1 := g.AddRadio("vcondition_group Checked", "От аванса")
	r2 := g.AddRadio("", "От подписания")
	r3 := g.AddRadio("", "От размещения заказа")
	r1.OnEvent("Click", set_order_condition.Bind(1))
	r2.OnEvent("Click", set_order_condition.Bind(2))
	r3.OnEvent("Click", set_order_condition.Bind(3))
	date_button := g.AddButton("w150 Default", "Рассчитать")
	date_button.OnEvent("Click", (*) => when_clicked())
	g.AddGroupBox("w150 h60", "Сроки поставки")
	result_text_1a := g.AddText("xp+5 yp+17 w70")
	result_text_1a.OnEvent("DoubleClick", copyText.Bind(result_text_1a))
	result_text_1b := g.AddText("x90 yp w70")
	result_text_1b.OnEvent("DoubleClick", copyText.Bind(result_text_1b))
	result_text_2a := g.AddText("x27 yp+20 w70")
	result_text_2a.OnEvent("DoubleClick", copyText.Bind(result_text_2a))
	result_text_2b := g.AddText("x90 yp w70")
	result_text_2b.OnEvent("DoubleClick", copyText.Bind(result_text_2b))
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
	add_button := g.AddButton("w150 Default", "Рассчитать")
	add_button.OnEvent("Click", (*) => add_click())
	g.AddGroupBox("w150 h130", "Новый срок поставки")
	; newDelivery := g.AddText("xp+5 yp+16 w125", "Новый срок поставки")
	new_date := g.AddText("xp+5 yp+20 w125")
	new_date.OnEvent("DoubleClick", copyText.Bind(new_date))
	new_date_weeks := g.AddText("w125")
	new_date_weeks.OnEvent("DoubleClick", copyText.Bind(new_date_weeks))

	add_click() {
		start_date := g.Submit(0).AddStart
		original_weeks := g.Submit(0).OrigTime
		new_buz_date := g.Submit(0).NewBuzDate
		fca_ddp_weeks := g.Submit(0).FcaDdp

		; Первоначальный срок поставки (long date)
		original_delivery_date := DateAdd(start_date, original_weeks*7, "Days")
		; Первоначальный срок поставки (формат)
		original_delivery_date_f := FormatTime(original_delivery_date, "dd.MM.yyyy")
		; Первоначальный срок поставки в неделях
		original_delivery_time_weeks := Integer(DateDiff(original_delivery_date, start_date, "Days") / 7)
		; Новый срок поставки (long date)
		new_ddp_date := DateAdd(new_buz_date, fca_ddp_weeks*7, "Days")
		; Новый срок поставки (формат)
		new_ddp_date_f := FormatTime(new_ddp_date, "dd.MM.yyyy")
		new_ddp_date_weeks := Ceil(DateDiff(new_ddp_date, start_date, "Days") / 7)
		; MsgBox("Первоначальный срок поставки: " originalDeliveryDateF " (" originalDeliveryTimeWeeks " нед.)`nНовый срок поставки: " newDDPDateF " (" newDDPDateWeeks " нед.)")

		ControlSetText(new_ddp_date_f, new_date)
		ControlSetText(Format("{} недель", new_ddp_date_weeks), new_date_weeks)
	}

	Tab3.UseTab("Сумма прописью")
	g.AddText(, "Сумма:")
	g.AddEdit("r1 w150 vSum")
	g.AddRadio("Checked vCurrencyRadioGroup", "RUB")
	g.AddRadio("yp", "EUR")
	g.AddRadio("x22 y120", "USD")
	g.AddRadio("yp", "CHF")
	g.AddText("x22 y144 w100", "НДС:")
	vat := [0, 7, 10, 12, 13, 15, 17, 18, 20]
	g.AddDropDownList("x55 y140 r9 Choose9 w117 vTax", vat)
	spell_button := g.AddButton("x22 y168 w150", "Превратить в текст")
	spell_button.OnEvent("Click", (*) => click_spell())
	sum_field := g.AddEdit("w150 r10 ReadOnly")

	click_spell() {
		spell_data := {
			digit_sum: RegExReplace(g.Submit(0).Sum, "[A-Za-z\s]*"),
			currency: g.Submit(0).CurrencyRadioGroup,
			vat: g.Submit(0).Tax
		}
		if spell_data.digit_sum = "" {
			ControlSetText("", sum_field)
			return
		}
		switch spell_data.currency{
			case 1: currency := "RUB"
			case 2: currency := "EUR"
			case 3: currency := "USD"
			case 4: currency := "USD"
		}
		spelt_sum := SpellSum(spell_data.digit_sum, currency, spell_data.vat)
		if spell_data.currency = 4 {
			spelt_sum := StrReplace(spelt_sum, "доллар", "франк")
			spelt_sum := StrReplace(spelt_sum, "франкы", "франки")
			spelt_sum := StrReplace(spelt_sum, "цент", "раппен")
			spelt_sum := StrReplace(spelt_sum, "USD", "CHF")
			spelt_sum := StrReplace(spelt_sum, " США")
		}
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
			1, "{1} с момента поступления авансового платежа на счет Поставщика.",
			2, "{1} с момента подписания спецификации уполномоченными представителями Поставщика и Покупателя.",
			3, "{1} с момента подписания заказа уполномоченными представителями Поставщика и Покупателя.",
		)
		if order_condition != 0
			A_Clipboard := Format(Conditions[order_condition], ControlGetText(ControlGetHwnd(Control)))
		else
			A_Clipboard := ControlGetText(ControlGetHwnd(Control))
		text_to_format := "Текст '{1}' скопирован!"
		Info(Format(text_to_format, A_Clipboard))
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

		costs := [price, convert_price(price, EUR_rate), convert_price(price, CHF_rate), convert_price(price, USD_rate), convert_price(price, CNY_rate)]
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

	switch_radio(calculationMode, *) {
		switch calculationMode {
			case "toOffer":
				Offer()
			case "toOrder":
				Order()
		}
		
		Offer() {
			ControlSetText("Дата КП", source_date)
			ControlSetText("+/- дней(EXW):", text_days)
			ControlSetText("+/- недель(DDP):", text_weeks)
			days_edit.Value := 1
			weeks_edit.Value := 11
			set_order_condition(g.Submit(0).condition_group)
			r1.Enabled := 1
			r2.Enabled := 1
			r3.Enabled := 1
		}
		Order() {
			ControlSetText("Дата заказа", source_date)
			ControlSetText("Недель от BUZ до клиента:", text_days)
			ControlSetText("Общий срок поставки:", text_weeks)
			days_edit.Value := 11
			weeks_edit.Value := 12
			set_order_condition(0)
			r1.Enabled := 0
			r2.Enabled := 0
			r3.Enabled := 0
		}

	}

	when_clicked() {
		is_offer := g.Submit(0).is_offer
		is_order := g.Submit(0).is_order

		; SetOrderCondition(g.Submit(0).condition_group)

		date.start_date := g.Submit(0).start_date
		date.weeks := g.Submit(0).weeks
		date.days := g.Submit(0).days
		calculate()

		calculate() {

			if date.weeks != ""
				number_of_weeks := Integer(date.weeks) * 7
			else
				number_of_weeks := 0
			if date.days != ""
				number_of_days := Integer(date.days)
			else
				number_of_days := 0

			if is_offer = true
				calculate_offer()
			else if is_order = true
				calculate_order()

			calculate_offer() {
				delivery_time := (Ceil(number_of_days / 5) * 7) + number_of_weeks
				calculated_date := DateAdd(date.start_date, delivery_time, "Days"), "dd.MM.yyyy" ;25.09.2023
				delivery_date := FormatTime(DateAdd(date.start_date, delivery_time, "Days"), "dd.MM.yyyy") ;25.09.2023
				delivery_weeks := Ceil(DateDiff(calculated_date, date.start_date, "Days") / 7)

				ControlSetText(delivery_date, result_text_1a)
				ControlSetText(delivery_weeks " недель", result_text_1b)
				ControlSetText("", result_text_2a)
				ControlSetText("", result_text_2b)
			}
			
			calculate_order() {
				weeks_to_russia := number_of_days * 7
				total_weeks := number_of_weeks
				calculation_ddp := DateAdd(date.start_date, total_weeks, "Days")
				calculation_fca := DateAdd(calculation_ddp, -weeks_to_russia, "Days")
				calculation_exw := DateAdd(calculation_ddp, -calculation_fca, "Days")
				FCA := FormatTime(calculation_fca, "dd.MM.yyyy")
				DDP := FormatTime(calculation_ddp, "dd.MM.yyyy")
				ControlSetText("Дата EXW:", result_text_1a)
				ControlSetText(FCA, result_text_1b)
				ControlSetText("Дата DDP:", result_text_2a)
				ControlSetText(DDP, result_text_2b)
			}

		}
	}


	button_event(*) {
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
		
		costs := [price, convert_price(price, EUR_rate), convert_price(price, CHF_rate), convert_price(price, USD_rate), convert_price(price, CNY_rate)]
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

	convert_price(price, currency) {
		return toComma(Format("{:.2f}", price / toDot(currency)))
	}
	g.Show()
}
bCalculator()