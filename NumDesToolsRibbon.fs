module MyRibbon

open System.Windows.Forms
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open ExcelDna.Integration
open ExcelDna.Integration.CustomUI
open System.Reflection
open System.IO   
open System.Text

// This defines a regular Excel macro (in Excel you can press Alt + F8, type in the name "showMessage", then click the Run button).
// For the ribbon, it will be run through the ExcelRibbon.RunTagMacro(...) helper, which run whatever macro is specified in the button tag attribute
// One advantage is that you can 
[<ExcelCommand>]
let showMessage () =
    XlCall.Excel(XlCall.xlcAlert, "Hello from a macro!") 
    |> ignore


// This type defines the ribbon interface. It is a public class that derives from ExcelRibbon
[<ComVisible(true)>]    // This attribute is only needed if there is an assembly-level [<assembly:ComVisible(false)>] attribute.
type public MyRibbon() =
    inherit ExcelRibbon()

    // The ribbon xml definition could also be placed in the .dna file
    // Remember to switch on the ExcelOption "Show add-in user interface errors" option (under the Advanced tab under General)

    override this.GetCustomUI(ribbonId) =
        let mutable ribbonXml = ""
        try
            // 调用资源读取函数
            ribbonXml <- this.GetRibbonXml("Ribbon.xml")

            // 调试模式下的字符串替换
#if DEBUG
            ribbonXml <- ribbonXml.Replace(
                "<tab id='MainTab' label='NumDesToolsF#' insertBeforeMso='TabHome'>",
                "<tab id='MainTab' label='N*D*T*DebugF#' insertBeforeMso='TabHome'>"
            )
            ribbonXml <- ribbonXml.Replace(
                "<tab id='SecondTab' label='NumDesToolsF#Plus' insertBeforeMso='TabHome'>",
                "<tab id='SecondTab' label='N*D*T*PlusDebugF#' insertBeforeMso='TabHome'>"
            )
#endif
        with
        | ex -> MessageBox.Show(ex.Message) |> ignore // F# 中忽略返回值使用 |>

        ribbonXml

    // 内部方法，用于从嵌入式资源读取 Ribbon XML
    member private this.GetRibbonXml(resourceName: string) =
        let assn = Assembly.GetExecutingAssembly()
        let resources = assn.GetManifestResourceNames()
        let mutable text = ""

        for resource in resources do
            if resource.EndsWith(resourceName) then
                use streamText = assn.GetManifestResourceStream(resource)
                if streamText <> null then
                    use reader = new StreamReader(streamText)
                    text <- reader.ReadToEnd()
                // F# 的 `use` 关键字会自动处理资源释放，无需手动 Close()

        text

    member this.OnButtonPressed (control:IRibbonControl) =
        MessageBox.Show "Hello from F#!" 
        |> ignore

    member this.OnDumpData (control:IRibbonControl) =
        let app = ExcelDnaUtil.Application :?> Application
        let cellA1 = app.Range("A1")
        cellA1.Value2 <- app.Version
        // could also replace the last line with
        //     cellA1.Value(XlRangeValueDataType.xlRangeValueDefault) <- app.Version 
        // but Range.Value is an indexer property, so it's a bit inconvenient
