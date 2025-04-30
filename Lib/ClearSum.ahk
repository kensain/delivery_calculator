#Requires AutoHotkey v2.0

ClearSum(Value) {

    if IsFloat(Value)
        Value := Format("{:.2f}", Value)

    res := RegExReplace(Value, "[^\d\.,]")

    if InStr(res, ".",, StrLen(res) - 2) {
        return StrReplace(res, ",")
    }
    if InStr(res, ",",, StrLen(res) - 2) {
        a := StrSplit(res, ",")
        return StrReplace(a[1], ".") "." a[2]
    }
    if IsInteger(res)
        return res ".00"

    return false
}