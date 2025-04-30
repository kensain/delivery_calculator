#Include ./ClearSum.ahk
/**
 * Возвращает число прописью через сайт {@link https://summa-propisyu.ru}.
 * @param {Integer|Float|String} Number Число для вывода прописью
 * @param {String} Currency Валюта суммы - "RUB", "USD" или "EUR"
 * @param {Integer|String} VAT Процент НДС: 0, 7, 10, 12, 13 (НДФЛ), 17, 18 или 20. "Без НДС" вообще уберёт упоминание НДС.
 * @param {String} Separator Разделитель дробной части - "," или ".".
 * @returns {String} Возвращает строку с суммой прописью.
 */
SpellSum(Number, Currency := "RUB", VAT := 20, Separator := ",") {
    switch Currency {
        case "RUB": _Currency := "0"
        case "USD": _Currency := "2"
        case "EUR": _Currency := "3"
        case "CNY": _Currency := "4"
        case "CHF": _Currency := "2"
    }
    switch Separator {
        case ".": _Separator := "0"
        case ",": _Separator := "1"
    }

    _Number := ClearSum(Number)
    _VAT := VAT = "Без НДС" ? String(0) : String(VAT)

    url := Format("https://summa-propisyu.ru/?summ={1}&vat={2}&val={3}&sep={4}", _Number, _VAT, _Currency, _Separator)

    response := ComObject("WinHttp.WinHttpRequest.5.1")
    response.Open("GET", url, false)
    response.Send()
    if response.Status != 200 {
        MsgBox("Request failed with status code " response.Status)
    } else {
        xml := response.ResponseText
        ; A_Clipboard := xml
        
        results_quantity := 11
        output := []
        loop 11 {
            needle := Format("(?<=id=result{1}>).*(?=</textarea>)", Format("{:02}", A_Index))
            RegExMatch(xml, needle, &match)
            output.Push(match[])
        }
        if Currency = "CHF" {
            OutputDebug("Is CHF`n")
            output[4] := StrReplace(output[4], " США")
            output[4] := StrReplace(output[4], "доллар", "франк")
            output[4] := StrReplace(output[4], "цент", "раппен")
        }
        return VAT = "Без НДС" ? StrReplace(output[4], ", НДС не облагается") : output[4]
    }
}

OutputDebug SpellSum("1344,44", "CHF")
/* class __SpellSum {

    VAT := Map(
        "Без НДС",  0,
        "7% НДС",   7,
        "10% НДС",  10,
        "12% НДС",  12,
        "17% НДС",  17,
        "18% НДС",  18,
        "20% НДС",  20,
        "13% НДФЛ", 13,
    )

    Currency := Map(
        "Рубль",    0,
        "Доллар",   2,
        "Евро",     3
    )

    Separator := Map(
        ".",    0,
        ",",    1
    )

    __New(Number, Currency := this.Currency["Евро"], VAT := this.VAT["20% НДС"], Sep := this.Separator[","]) {
        url := Format("https://summa-propisyu.ru/?summ={1}&vat={2}&val={3}&sep={4}", Number, VAT, Currency, Sep)

        response := ComObject("WinHttp.WinHttpRequest.5.1")
        response.Open("GET", url, false)
        response.Send()
        if response.Status != 200 {
            MsgBox("Request failed with status code " response.Status)
        } else {
            xml := response.ResponseText
            results_quantity := 11
            FileOpen("output.xml", "w", "UTF-16").Write(xml)
            output := Map()
            loop 11 {
                needle := Format("(?<=id=result{1}>).*(?=</textarea>)", A_Index)
                RegExMatch(xml, needle, &match)
                output.Set(A_Index, match[])
            }
            return output[4]
        }
    }

} */