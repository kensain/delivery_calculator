/**
 * Get the current exchange rate from the Central Bank if Russia. V2.
 * @param {String} Date Optional date for the exchange rate (05.05.2023 or 05/05/2023 format). If omitted, the the most recent official date will be used.
 * @returns {Object} Returns an object with the following ```Currency``` (Map) and ```Date``` (String). Example output: ```CBR().Currency["EUR"]```
 */
CBR2(Date?) {
	WinHttp := ComObject("WinHttp.WinHttpRequest.5.1")
	url := "http://www.cbr.ru/scripts/XML_daily.asp"
	if IsSet(Date)
		url .= "?date_req=" . StrReplace(Date, ".", "/")
	WinHttp.Open("GET", url)
	WinHttp.Send()
	res := WinHttp.ResponseText

	Currencies := Map()

	doc := loadXML(res)
	for valute in doc.getElementsByTagName("Valute") {
		Currencies.Set(
			valute.selectSingleNode("CharCode").Text,
			valute.selectSingleNode("Value").Text
		)
	}

	OutputObj := {
		Currency: Currencies,
		Date: doc.selectSingleNode("/ValCurs/@Date").Text
	}

	return OutputObj
	
	loadXML(data) {
	  o := ComObject("MSXML2.DOMDocument.6.0")
	  o.async := false
	  o.loadXML(data)
	  return o
	}
}
