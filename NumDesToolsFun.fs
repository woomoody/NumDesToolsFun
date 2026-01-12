namespace NumDesToolsFun

open System
open System.IO
open System.Text
open System.Text.RegularExpressions
open Newtonsoft.Json
open ExcelDna.Integration
open System.Collections.Generic
open System.Linq

module NumDesToolsFun = 

    [<ExcelFunction(Description = "My first .NET function")>]
    let SayHello name = "Hello " + name

    [<ExcelFunction(Category="UDF-组装字符串", IsVolatile=true, IsMacroType=true,Description="拼接Range数据")>]
    let CreatValueToArrayFun
        ([<ExcelArgument(Description="单元格范围")>] rangeObj: obj[,])
        ([<ExcelArgument(Description="默认值范围")>] rangeObjDef: obj[,])
        ([<ExcelArgument(Description="分隔符")>] delimiter: string)
        ([<ExcelArgument(Description="过滤值")>] ignoreValue: string)  =

        // 添加类型注解，使用完整的类型名称
        let isInvalidCell (value: obj) =
            match value with
            | :? ExcelDna.Integration.ExcelEmpty as _ -> true
            | :? ExcelDna.Integration.ExcelError as _ -> true
            | cell when cell.ToString() = ignoreValue -> true
            | _ -> false

        let (start, mid, finish) =
            let delim = if System.String.IsNullOrEmpty(delimiter) then "[,]" else delimiter
            match delim.ToCharArray() |> Array.toList with
            | [s; m; f] -> (string s, string m, string f)
            | m::_ -> ("", string m, "")
            | [] -> ("", ",", "")

        let useDefault = 
            let firstCell = rangeObjDef.[0, 0]
            match firstCell with
            | :? ExcelDna.Integration.ExcelMissing -> false
            | _ -> true

        // 使用数组推导式避免类型推断问题
        let allCells = 
            [| for i in 0 .. Array2D.length1 rangeObj - 1 do
                for j in 0 .. Array2D.length2 rangeObj - 1 do
                    if not (isInvalidCell rangeObj.[i, j]) then
                        if useDefault then
                            yield rangeObjDef.[i, j].ToString()
                        else
                            yield rangeObj.[i, j].ToString() |]

        if allCells.Length = 0 then
            ""
        else
            let result = String.concat mid allCells
            start + result + finish