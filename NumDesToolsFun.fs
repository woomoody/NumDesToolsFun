namespace NumDesToolsFun

module NumDesToolsFun = 
    open ExcelDna.Integration

    [<ExcelFunction(Description = "My first .NET function")>]
    let SayHello name = "Hello " + name
