/**
 * Возвращает число прописью через сайт https://summa-propisyu.ru.
 * @param Number Число для вывода прописью
 * @param {string} Currency Валюта суммы - "RUB", "USD" или "EUR"
 * @param {number} VAT Процент НДС: 0, 7, 10, 12, 13 (НДФЛ), 17, 18 или 20.
 * @param {string} Separator Разделитель дробной части - "," или ".".
 * @returns {string} Возвращает строку с суммой прописью.
 */
SpellSum(Number, Currency := "RUB", VAT := 20, Separator := ",") {
    switch Currency {
        case "RUB": _Currency := "0"
        case "USD": _Currency := "2"
        case "EUR": _Currency := "3"
    }
    switch Separator {
        case ".": _Separator := "0"
        case ",": _Separator := "1"
    }
    VAT := String(VAT)

    url := Format("https://summa-propisyu.ru/?summ={1}&vat={2}&val={3}&sep={4}", Number, VAT, _Currency, _Separator)

    response := ComObject("WinHttp.WinHttpRequest.5.1")
    response.Open("GET", url, false)
    response.Send()
    if response.Status != 200 {
        MsgBox("Request failed with status code " response.Status)
    } else {
        xml := response.ResponseText
        results_quantity := 11
        ; FileOpen("output.xml", "w", "UTF-16").Write(xml)
        output := Map()
        loop 11 {
            needle := Format("(?<=id=result{1}>).*(?=</textarea>)", A_Index)
            RegExMatch(xml, needle, &match)
            output.Set(A_Index, match[])
        }
        return output[4]
    }
}

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