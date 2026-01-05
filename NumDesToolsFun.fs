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
        ([<ExcelArgument(Description="过滤值")>] ignoreValue: string) =
        
        let mutable result = ""
        let delimiter = if String.IsNullOrEmpty(delimiter) then "[,]" else delimiter
        let delimiterList = delimiter.ToCharArray() |> Array.map string
        
        let startDelimiter, midDelimiter, endDelimiter = 
            if delimiterList.Length = 3 then
                delimiterList.[0], delimiterList.[1], delimiterList.[2]
            else
                "", delimiterList.[0], ""
        
        let rows = rangeObj.GetLength(0)
        let cols = rangeObj.GetLength(1)
        
        for row in 0 .. rows - 1 do
            for col in 0 .. cols - 1 do
                let item = rangeObj.[row, col]
                match item with
                | :? ExcelEmpty -> ()
                | _ when item.ToString() = ignoreValue -> ()
                | :? ExcelError -> ()
                | _ ->
                    if rangeObjDef.[0, 0] :? ExcelMissing then
                        result <- result + item.ToString() + midDelimiter
                    else
                        result <- result + rangeObjDef.[row, col].ToString() + midDelimiter
        
        if not (String.IsNullOrEmpty(result)) then
            startDelimiter + result.Substring(0, result.Length - 1) + endDelimiter
        else
            ""