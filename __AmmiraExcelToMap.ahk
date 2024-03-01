#Include <Json>

F7::{
    xl := ComObjActive("Excel.Application")
    TARIFFS := 8
    result := Map()
    for city in xl.Selection {
        rates := []
        loop TARIFFS {
            rates.Push(city.Offset(0, 0 + A_Index).Value)
        }
        result.Set(city.Text, rates)
    }
    A_Clipboard := JSON.stringify(result)
    MsgBox("Done!")
    ExitApp()
}