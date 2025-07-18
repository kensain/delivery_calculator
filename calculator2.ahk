#SingleInstance Force
#Requires AutoHotkey v2.0-a
#Include ./Lib/tariffs_2025.ahk
#Include ./Lib/DeliveryCosts.ahk
#Include <Info>
#Include <CBR2>
#Include <SpellSum>

bCalculator() {

	TraySetIcon("icon.ico")

	OrderCondition := 1
	SetOrderCondition(Int) {
		OrderCondition := Int
		Info(Int)
	}

	g := Gui()
	g.Opt("AlwaysOnTop ToolWindow")
	g.OnEvent("Escape", (*) => __CloseGui())
	g.OnEvent("Close", (*) => __CloseGui())
	g.Title := "Калькулятор доставки 2025"

	Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост", "Аммира", "Даты", "Сумма прописью", "Международная доставка"])
	; [Оставить для возможного возврата вкладки "Для ДС" в будущем]{
		; Tab3 := g.Add("Tab3", "vTab3", ["Гарантпост", "Аммира", "Даты", "Для ДС", "Сумма прописью"])
	; }


	BLOCKED_REGIONS := Map(
		"Белгородская область",	"Белгородскую область считаем через Аммиру",
		"Белгород",				"Белгородскую область считаем через Аммиру",
		"Брянская область",		"Брянскую область считаем через Аммиру",
		"Брянск",				"Брянскую область считаем через Аммиру",
		"Воронежская область",	"Воронежскую область считаем через Аммиру",
		"Воронеж",				"Воронежскую область считаем через Аммиру",
		"Курская область",		"Курскую область считаем через Аммиру",
		"Курск",				"Курскую область считаем через Аммиру",
		"Краснодарский край",	"Краснодарскую область считаем через Аммиру",
		"Краснодар",			"Краснодарскую область считаем через Аммиру",
		"Ростовская область",	"Ростовскую область считаем через Аммиру",
		"Ростов-на-Дону",		"Ростовскую область считаем через Аммиру",
		"Азов", 				"Ростовская область",
		"Аксай", 				"Ростовская область",
		"Алексеевка", 			"Белгородская область",
		"Батайск", 				"Ростовская область",
		"Белая Глина", 			"Краснодарский край",
		"Затонский", 			"Ростовская область",
		"Кореновск", 			"Краснодарский край",
		"Краснодар", 			"Краснодарский край",
		"Севск", 				"Брянская область",
		"Славянск-на-Кубани", 	"Краснодарский край",
		"Советское", 			"Белгородская область?",
		"Старый Оскол", 		"Белгородская область",
		"Таганрог", 			"Ростовская область",
		"Тамань", 				"Краснодарский край",
		"Тимашевск", 			"Краснодарский край",
		"Шебекино", 			"Белгородская область"
	)

	garantpostList := []
	for each in GARANTPOST
		garantpostList.Push(each[1])
	g.AddComboBox("r18 vGarantpostChoice Choose1 w150 AltSubmit Simple", garantpostList)
	g.AddText(, "Вес в КГ:")
	g.AddEdit("r1 vWeight w150 Number")
	g.AddButton("vButtonGarant Default w150 Disabled", "OK")
	g["ButtonGarant"].OnEvent("Click", ClickEventGarantPost)
	g["Weight"].OnEvent("Change", (*) => __CheckState("ButtonGarant", "Weight"))
	Tab3.UseTab()
	g.AddCheckbox("vBlockCertainRegions Checked", "Исключить области")
	g.AddText(, "Курсы на")
	CalendarDate := "Choose" . FormatTime(A_Now, "yyyyMMdd")
	g.AddDateTime("yp w108 vCBRDate " . CalendarDate, "dd.MM.yyyy").OnEvent("Change", (*) => UpdateRates())

	; [Currencies grid] {
	g.AddText("vtEUR r1 w22 xm", "EUR")
	g.AddEdit("veEUR ReadOnly r1 w55 yp")
	g.AddText("vtCHF r1 w22 yp", "CHF")
	g.AddEdit("veCHF ReadOnly r1 w55 yp")
	g.AddText("vtUSD r1 w22 xm", "USD")
	g.AddEdit("veUSD ReadOnly r1 w55 yp")
	g.AddText("vtCNY r1 w22 yp", "CNY")
	g.AddEdit("veCNY ReadOnly r1 w55 yp")
	g["tEUR"].OnEvent("DoubleClick", (*) => Info(g["eEUR"].Text))
	g["tCHF"].OnEvent("DoubleClick", (*) => Info(g["eCHF"].Text))
	g["tUSD"].OnEvent("DoubleClick", (*) => Info(g["eUSD"].Text))
	g["tCNY"].OnEvent("DoubleClick", (*) => Info(g["eCNY"].Text))
	g.AddLink("xm r1 w180 vCBRLink")
	; }
	
	UpdateRates()
	
	; [Вкладка Аммира] {
	Tab3.UseTab("Аммира")
	AmmiraCities := []
	for key, value in AMMIRA
		AmmiraCities.Push(key)
	g.AddComboBox("r18 vAmmiraChoice Choose1 w150 Simple", AmmiraCities)
	g.AddText(, "Вес в КГ:")
	g.AddEdit("r1 vWeightAmmira w150 Number")
	g["WeightAmmira"].OnEvent("Change", (*) => __CheckState("ButtonAmmira", "WeightAmmira"))
	g.AddButton("vButtonAmmira Default w150 Disabled", "OK").OnEvent("Click", ClickEventAmmira)
	; }

	; [Вкладка "Даты"] {
	Tab3.UseTab("Даты")
	g.AddRadio("vIsOfferGroup Checked1", "КП клиенту").OnEvent("Click", SwitchRadio.Bind("toOffer"))
	g.AddRadio("Checked0", "Размещение заказа").OnEvent("Click", SwitchRadio.Bind("toOrder"))
	g.AddText("vSourceDate r1", "Дата КП:")
	g.AddDateTime("yp-3 x75 vStartDate w97", "dd.MM.yyyy")
	g.AddText("vDaysText x22 yp+25 r1 w150", "+/- дней(EXW):")
	g.AddEdit("w150")
	g.AddUpDown("vDays Range0-180", 1).OnEvent("Change", (*) => CheckStateDate())
	g.AddText("r1 w150 vDDPweeks", "+/- недель(DDP):")
	g.AddEdit("w150")
	g.AddUpDown("vWeeks Range0-180", 11).OnEvent("Change", (*) => CheckStateDate())
	r1 := g.AddRadio("vConditionGroup Checked", "От аванса")
	r2 := g.AddRadio("", "От подписания")
	r3 := g.AddRadio("", "От размещения заказа")
	r1.OnEvent("Click", (*) => SetOrderCondition(1))
	r2.OnEvent("Click", (*) => SetOrderCondition(2))
	r3.OnEvent("Click", (*) => SetOrderCondition(3))
	g.AddButton("vDateButton w150 Default", "Рассчитать").OnEvent("Click", (*) => ClickEventDate())
	g.AddGroupBox("w150 h60", "Сроки поставки")
	g.AddText("vResultText1a xp+5 yp+17 w70")
	g.AddText("vResultText1b x90 yp w70").OnEvent("DoubleClick", (*) => CopyText("ResultText1b"))
	g.AddText("vResultText2a x27 yp+20 w70")
	g.AddText("vResultText2b x90 yp w70").OnEvent("DoubleClick", (*) => CopyText("ResultText2b"))
	; }
	; }
; [Оставить на будущее] {
; 	Tab3.UseTab("Для ДС")
; 	g.AddText(, "Точка отсчёта")
; 	g.AddDateTime("vAddStart w150", "dd.MM.yyyy")
; 	g.AddText(, "Изначальный срок (нед.)")
; 	g.AddEdit("w150")
; 	g.AddUpDown("vOrigTime Range0-180", 1)
; 	g.AddText(, "Новая дата готовности")
; 	g.AddDateTime("vNewBuzDate w150", "dd.MM.yyyy")
; 	g.AddText(, "Недель от BUZ до клиента:")
; 	g.AddEdit("w150")
; 	g.AddUpDown("vFcaDdp Range0-180", 11)
; 	g.AddButton("w150 Default", "Рассчитать").OnEvent("Click", (*) => add_click())
; 	g.AddGroupBox("w150 h130", "Новый срок поставки")
; 	g.AddText("vNewDate xp+5 yp+20 w125")
; 	g["NewDate"].OnEvent("DoubleClick", CopyText.Bind("NewDate"))
; 	g.AddText("vNewDateWeeks w125")
; 	g["NewDateWeeks"].OnEvent("DoubleClick", CopyText.Bind("NewDateWeeks"))
	
; 	add_click() {
; 		StartDate := g.Submit(0).AddStart
; 		OriginalWeeks := g.Submit(0).OrigTime
; 		NewBUZDate := g.Submit(0).NewBuzDate
; 		FCADDPWeeks := g.Submit(0).FcaDdp
		
; 		; Первоначальный срок поставки (long date)
; 		OriginalDeliveryDate := DateAdd(StartDate, OriginalWeeks*7, "Days")
; 		; Первоначальный срок поставки (формат)
; 		fOriginalDeliveryDate := FormatTime(OriginalDeliveryDate, "dd.MM.yyyy")
; 		; Первоначальный срок поставки в неделях
; 		OriginalDeliveryTimeWeeks := Integer(DateDiff(OriginalDeliveryDate, StartDate, "Days") / 7)
; 		; Новый срок поставки (long date)
; 		NewDDPDate := DateAdd(NewBUZDate, FCADDPWeeks*7, "Days")
; 		; Новый срок поставки (формат)
; 		fNewDDPDate := FormatTime(NewDDPDate, "dd.MM.yyyy")
; 		NewDDPDateWeeks := Ceil(DateDiff(NewDDPDate, StartDate, "Days") / 7)
; 		; MsgBox("Первоначальный срок поставки: " originalDeliveryDateF " (" originalDeliveryTimeWeeks " нед.)`nНовый срок поставки: " newDDPDateF " (" newDDPDateWeeks " нед.)")
		
; 		__UpdateText("NewDate", fNewDDPDate)
; 		__UpdateText("NewDateWeeks", Format("{} недель", NewDDPDateWeeks))
; 	}
; 	}
	Tab3.UseTab("Сумма прописью")
	g.AddText(, "Сумма:")
	g.AddEdit("r1 w150 vInputSum")
	g.AddRadio("Checked vCurrencyRadioGroup", "RUB")
	g.AddRadio("yp", "EUR")
	g.AddRadio("x22 yp+25", "USD")
	g.AddRadio("yp", "CHF")
	g.AddText("x22 yp+25 w100", "НДС:")
	vat := [0, 7, 10, 12, 13, 15, 17, 18, 20]
	g.AddDropDownList("x55 yp r9 Choose9 w117 vTax", vat)
	g.AddButton("x22 yp+25 w150", "Превратить в текст").OnEvent("Click", (*) => ClickEventSpell())
	g.AddEdit("vOutputSum w150 r10 ReadOnly")
	
	ClickEventSpell() {
		SpellObj := {
			DigitSum: RegExReplace(g.Submit(0).InputSum, "[A-Za-z\s]*"),
			Currency: g.Submit(0).CurrencyRadioGroup,
			VAT: g.Submit(0).Tax
		}
		if SpellObj.DigitSum = "" {
			UpdateText("OutputSum")
			return
		}
		switch SpellObj.Currency{
			case 1: Currency := "RUB"
			case 2: Currency := "EUR"
			case 3: Currency := "USD"
			case 4: Currency := "USD"
		}
		SpeltSum := SpellSum(SpellObj.DigitSum, Currency, SpellObj.VAT)
		; Т.к. сайт не поддерживает франки, пришлось делать через переименование долларов:
		if SpellObj.Currency = 4 { 
			SpeltSum := StrReplace(SpeltSum, "доллар", "франк")
			SpeltSum := StrReplace(SpeltSum, "франкы", "франки")
			SpeltSum := StrReplace(SpeltSum, "цент", "раппен")
			SpeltSum := StrReplace(SpeltSum, "USD", "CHF")
			SpeltSum := StrReplace(SpeltSum, " США")
		}
		UpdateText("OutputSum", SpeltSum)
	}

	; [Вкладка "Международная доставка"] {
	Tab3.UseTab("Международная доставка")
	g.AddText("r1", "Стоимость доставки (евро):")
	g.AddEdit("w150 vDeliveryCostEUR", 5600)
	g.AddText("r1", "Наценка (%):")
	g.AddEdit("w15 vMarkup", "3")
	g.AddButton("vCalculateDeliveryButton w150 Default", "Пересчитать в рубли").OnEvent("Click", (*) => ClickCalculateDeliveryCost())
	g.AddEdit("w150 vDeliveryCostRUB")

	ClickCalculateDeliveryCost() {
		; Initiate
		Input := Float(RegExReplace(g["DeliveryCostEUR"].Text, "[^.,\d]"))
		Percentage := Float(g["Markup"].Text)
		CurrencyRate := Float(__ToDot(g["eEUR"].Text))
		; Calculate formula: (input * currency rate) * (1 + (Percentage / 100))
		InputRUB := Input * CurrencyRate
		PercentageIncrease := InputRUB * (1 + (Percentage / 100))
		; Update GUI
		g["DeliveryCostRUB"].Text := __ToComma(Round(PercentageIncrease, 2))
	}


	; g.AddRadio("vIsOfferGroup Checked1", "КП клиенту").OnEvent("Click", SwitchRadio.Bind("toOffer"))
	; g.AddRadio("Checked0", "Размещение заказа").OnEvent("Click", SwitchRadio.Bind("toOrder"))
	; g.AddText("vSourceDate r1", "Дата КП:")
	; g.AddDateTime("yp-3 x75 vStartDate w97", "dd.MM.yyyy")
	; g.AddText("vDaysText x22 y114 r1 w150", "+/- дней(EXW):")
	; g.AddEdit("w150")
	; g.AddUpDown("vDays Range0-180", 1).OnEvent("Change", (*) => CheckStateDate())
	; g.AddText("r1 w150 vDDPweeks", "+/- недель(DDP):")
	; g.AddEdit("w150")
	; g.AddUpDown("vWeeks Range0-180", 11).OnEvent("Change", (*) => CheckStateDate())
	; r1 := g.AddRadio("vConditionGroup Checked", "От аванса")
	; r2 := g.AddRadio("", "От подписания")
	; r3 := g.AddRadio("", "От размещения заказа")
	; r1.OnEvent("Click", (*) => SetOrderCondition(1))
	; r2.OnEvent("Click", (*) => SetOrderCondition(2))
	; r3.OnEvent("Click", (*) => SetOrderCondition(3))
	; g.AddButton("vDateButton w150 Default", "Рассчитать").OnEvent("Click", (*) => ClickEventDate())
	; g.AddGroupBox("w150 h60", "Сроки поставки")
	; g.AddText("vResultText1a xp+5 yp+17 w70")
	; g.AddText("vResultText1b x90 yp w70").OnEvent("DoubleClick", (*) => CopyText("ResultText1b"))
	; g.AddText("vResultText2a x27 yp+20 w70")
	; g.AddText("vResultText2b x90 yp w70").OnEvent("DoubleClick", (*) => CopyText("ResultText2b"))
	
	Tab3.UseTab()
	
	ClickEventAmmira(*) {
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

		if g.Submit().BlockCertainRegions {
			if BLOCKED_REGIONS.Has(destination) {
				if MsgBox(BLOCKED_REGIONS[destination] "`n`nПодготовить письмо?",, "0x30 0x4") = "Yes"
					WriteToAmmira(Weight)
			} else {
				CalculateAmmira()
			}
		} else {
			CalculateAmmira()
		}
		CalculateAmmira() {
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
			
			price := AMMIRA[destination][tariff]
			
			costs := [price, __FormatPrice(price, g["eEUR"].Value), __FormatPrice(price, g["eCHF"].Value), __FormatPrice(price, g["eUSD"].Value), __FormatPrice(price, g["eCNY"].Value)]
			title := Format("Стоимость доставки {1} кг в {2}", weight, destination)
			DeliveryCosts(title, costs*)
		}
	}
	
	__CheckState(vButton, vEdit) {
		g[vButton].Enabled := g[vEdit].Value != "" ? 1 : 0
	}
	CheckStateDate(*) {
		g["DateButton"].Enabled := (g["Weeks"].Value != "" or g["Days"].Value != "") ? 1 : 0
	}
	
	SwitchRadio(CalculationMode, *) {
		switch CalculationMode {
			case "toOffer":
				Offer()
				case "toOrder":
					Order()
				}
				
		Offer() {
			UpdateText("SourceDate", "Дата КП")
			UpdateText("DaysText", "+/- дней(EXW):")
			UpdateText("DDPWeeks", "+/- недель(DDP):")
			UpdateText("Days", 1)
			UpdateText("Weeks", 11)
			SetOrderCondition(g.Submit(0).ConditionGroup)
			r1.Enabled := 1
			r2.Enabled := 1
			r3.Enabled := 1
		}
		Order() {
			UpdateText("SourceDate", "Дата заказа")
			UpdateText("DaysText", "Недель от BUZ до клиента:")
			UpdateText("DDPWeeks", "Общий срок поставки:")
			UpdateText("Days", 11)
			UpdateText("Weeks", 12)
			SetOrderCondition(0)
			r1.Enabled := 0
			r2.Enabled := 0
			r3.Enabled := 0
		}

	}
	
	ClickEventDate() {

		Date := {
			StartDate: g.Submit(0).StartDate,
			Weeks: g.Submit(0).weeks,
			Days: g.Submit(0).days
		}
		
		NumberOfWeeks := Date.Weeks != "" ? Integer(Date.Weeks) * 7 : 0
		NumberOfDays := Date.Days != "" ? Integer(Date.Days) : 0
		
		if g.Submit(0).IsOfferGroup = 1 {
			DeliveryTime := (Ceil(NumberOfDays / 5) * 7) + NumberOfWeeks
			CalculatedDate := DateAdd(Date.StartDate, DeliveryTime, "Days"), "dd.MM.yyyy"
			DeliveryDate := FormatTime(DateAdd(Date.StartDate, DeliveryTime, "Days"), "dd.MM.yyyy")
			DeliveryWeeks := Ceil(DateDiff(CalculatedDate, Date.StartDate, "Days") / 7)
			
			UpdateText("ResultText1a", DeliveryDate)
			UpdateText("ResultText1b", DeliveryWeeks " недель")
			UpdateText("ResultText2a")
			UpdateText("ResultText2b")
		} else {
			WeeksToRussia := NumberOfDays * 7
			WeeksTotal := NumberOfWeeks
			CalculationDDP := DateAdd(Date.StartDate, WeeksTotal, "Days")
			CalculationFCA := DateAdd(CalculationDDP, -WeeksToRussia, "Days")
			CalculationEXW := DateAdd(CalculationDDP, -CalculationFCA, "Days")
			FCA := FormatTime(CalculationFCA, "dd.MM.yyyy")
			DDP := FormatTime(CalculationDDP, "dd.MM.yyyy")
			UpdateText("ResultText1a", "Дата EXW:")
			UpdateText("ResultText1b", FCA)
			UpdateText("ResultText2a", "Дата DDP:")
			UpdateText("ResultText2b", DDP)
		}

	}

	WriteToAmmira(Weight := "") {
		g := Gui()
		g.Title := "Письмо в Аммиру"
		g.Opt("ToolWindow AlwaysOnTop")
		g.AddText(, "Габариты (размер или объём)")
		g.AddEdit("r1 vDims")
		g.AddText(, "Вес (кг):")
		g.AddEdit("r1 vWeight w40", Weight)
		g.AddText(, "Клиент:")
		g.AddEdit("r1 vCustomer")
		g.AddText(, "Адрес клиента:")
		g.AddEdit("r3 vAddress")
		button := g.AddButton("Default w80", "OK")
		g.OnEvent("Escape", (*) => CloseGui())
		g.OnEvent("Close", (*) => CloseGui())
		button.OnEvent("Click", (*) => WhenClicked())
		g.Show()
		Data := {
			Dimensions: "",
			CustomerName: "",
			CustomerAddress: "",
			Weight: ""
		}
	
		WhenClicked() {
			gSub := g.Submit()
			Data.Dimensions := gSub.Dims = "" ? "____" : gSub.Dims
			Data.CustomerName := gSub.Customer = "" ? "____" : gSub.Customer
			Data.CustomerAddress := gSub.Address = "" ? "____" : gSub.Address
			Data.Weight := gSub.Weight = "" ? "____" : gSub.Weight
	
			WriteMail()
		}
	
		CloseGui() {
			g.Destroy()
			Exit()
		}
	
		WriteMail() {
			; signature := "`n`nС уважением,`n`nПортнов Максим`nМенеджер по работе с клиентами`n`nООО «Бюлер Сервис»`nул. Отрадная, д. 2Б, стр. 1,`n127273 Москва, Россия`nТел.:  +7 495 139 34 00 (доб.162)`nМоб.:  +7 916 420 79 60`n`nmaxim.portnov@buhlergroup.com`nwww.buhlergroup.com"
			
			try {
				Outlook := ComObjActive("Outlook.Application")
			} catch Error as e {
				MsgBox("Ошибка: " e.Message "`nВозможно, Outlook не запущен.")
				Exit()
			}
			
			email := Outlook.CreateItem(0)
			email.BodyFormat := 1 ; olFormatHTML
			email.Subject := "Расчёт доставки // " Data.CustomerName
			mailText := "Здравствуйте, Фёдор,`n`nпрошу рассчитать стоимость доставки до клиента " Data.CustomerName ".`n`nАдрес доставки: " Data.CustomerAddress "."
			; if Data.dimensions != ""
				mailText .= "`nГабариты: " Data.Dimensions ", вес: " Data.Weight " кг."
			; mailText .= signature
			email.Body := mailText
			email.To := "tk.ammira@gmail.com"
			;email.CC := "bgoodman@vip-logistics.ru; nikolai.lovyagin@buhlergroup.com"
			email.Display()
			Outlook := ""
		}
	}

	ClickEventGarantPost(*) {
		Destination := g.Submit().GarantpostChoice
		Weight := g.Submit().Weight

		if g.Submit().BlockCertainRegions {
			if BLOCKED_REGIONS.Has(GARANTPOST[Destination][1]) {
				if MsgBox(BLOCKED_REGIONS[GARANTPOST[Destination][1]] "`n`nПодготовить письмо?",, "0x30 0x4") = "Yes" {
					WriteToAmmira(Weight)
				} else {
					CalculateGarantPost()
				}
			} else {
				CalculateGarantPost()
			}
		} else {
			CalculateGarantPost()
		}
		
		CalculateGarantPost() {
			; Conditions:
			if Weight <= 0.1
				Tariff := 2
			else if Weight <= 0.5
				Tariff := 3
			else if Weight <= 1
				Tariff := 4
			else if Weight > 1 {
				Tariff := 4
				Markup := Destination = 2 and Weight > 32 ? "****" : 5 ; если СПб и больше 32 кг
			}
			
			; Calculate price
			if !IsSet(Markup)
				Price := GARANTPOST[Destination][Tariff]
			else {
				if Markup = "****"
					Price := "****"
			else
				Price := GARANTPOST[Destination][Tariff] + ((Weight - 1) * GARANTPOST[Destination][Markup])
			}
			
			if Price != "****" {
				Costs := [
					Price,
					__FormatPrice(Price, g["eEUR"].Value),
					__FormatPrice(Price, g["eCHF"].Value),
					__FormatPrice(Price, g["eUSD"].Value),
					__FormatPrice(Price, g["eCNY"].Value)
				]
				Title := Format("Стоимость доставки {1} кг в {2}", Weight, GARANTPOST[Destination][1])
				DeliveryCosts(Title, Costs*)
			} else {
				Result := MsgBox(Format("К сожалению, для отправлений в {1} свыше 32 кг действует специальный тариф с применением регрессивной шкалы за каждый следующий кг, поэтому надо пересчитывать на сайте. `nОткрыть калькулятор на сайте?", GARANTPOST[Destination][1]), Format("Отправка в {}", GARANTPOST[Destination][1]), "YesNo")
				if Result = "Yes"
					Run("https://garantpost.ru/tools")
				else
					return
			}
		}
	}

	__CloseGui() {
		g.Destroy()
		Exit()
	}

	__ToDot(NumberWithComma) {
		return StrReplace(NumberWithComma, ",", ".")
	}

	__ToComma(NumberWithDot) {
		return StrReplace(NumberWithDot, ".", ",")
	}

	__FormatPrice(Price, Currency) {
		return __ToComma(Format("{:.2f}", Price / __ToDot(Currency)))
	}
	
	UpdateText(vControlName, NewText := "") {
		g[vControlName].Value := NewText
	}
	
	__ConvertDate(date) {
		input := StrSplit(date, ".")
		output := "Choose" . input[3] . input[2] . input[1]
		return output
	}

	UpdateRates(date?) {
		if !IsSet(date)
			date := FormatTime(g.Submit(false).CBRDate, "dd.MM.yyyy")
		NewRates := CBR2(date)
		date_rate := NewRates.date
		g["eEUR"].Value := NewRates.Currency["EUR"]
		g["eCHF"].Value := NewRates.Currency["CHF"]
		g["eUSD"].Text := NewRates.Currency["USD"]
		g["eCNY"].Text := NewRates.Currency["CNY"]
		g["CBRDate"].Value := StrSplit(date_rate, ".")[3] . StrSplit(date_rate, ".")[2] . StrSplit(date_rate, ".")[1]
		g["CBRLink"].Text := Format('Проверить курс на <a href="https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={1}">сайте</a> ЦБ РФ', date_rate)
	}
		
	/**
	 * 
	 * @param {String} Mode 'Date' or 'Rate'
	 * @param {String} textToCopy Text to copy
	*/
	CopyText(vControlName) {
		if g[vControlName].Text = ""
			return
		Conditions := Map(
			1, "{1} с момента поступления авансового платежа на счет Поставщика.",
			2, "{1} с момента подписания спецификации уполномоченными представителями Поставщика и Покупателя.",
			3, "{1} с момента подписания заказа уполномоченными представителями Поставщика и Покупателя.",
		)
		A_Clipboard := OrderCondition != 0 ? Format(Conditions[OrderCondition], g[vControlName].Text) : g[vControlName].Text
		Info(Format("Текст '{1}' скопирован!", A_Clipboard))
	}

	g.Show()
}