// BenchmarkAllCubeMeasures.fsx

let serverName = "[server name]" 
let asDbName = "[database name]"
let cubeName = "[Cube Name]" 
let numTimings = 2

#I @"C:\Program Files (x86)\Microsoft.NET\ADOMD.NET\140"
#I @"C:\Program Files (x86)\Microsoft.NET\ADOMD.NET\130"
#I @"C:\Program Files (x86)\Microsoft.NET\ADOMD.NET\120"
#r "Microsoft.AnalysisServices.AdomdClient.dll"
open System
open System.Data // CommandType
open System.IO
open System.Xml
open System.Diagnostics
open Microsoft.AnalysisServices.AdomdClient
open System.Collections.Generic


let adoConnStr = sprintf "Data Source=%s;Catalog=%s;Cube=%s" serverName asDbName cubeName

let adomdConn = new AdomdConnection(adoConnStr)
adomdConn.Open()

open System.Collections.Generic

let clearCache () = 
    let cmd = adomdConn.CreateCommand()
    cmd.CommandText <- sprintf @"<ClearCache xmlns=""http://schemas.microsoft.com/analysisservices/2003/engine"">
        <Object><DatabaseID>%s</DatabaseID></Object></ClearCache>" asDbName 
    // A more agressive cache clear procedure is in the Analysis Services Stored Procedure project:
    //cmd.CommandText <- "call ASSP.ClearAllCaches()" 
    cmd.CommandType <- CommandType.Text
    cmd.ExecuteNonQuery() |> ignore

let timeMdx (qStr:string)  = 
    let cmd = adomdConn.CreateCommand()
    cmd.CommandText <- qStr
    cmd.CommandType <- CommandType.Text
    let stopWatch = new Stopwatch()
    stopWatch.Start()
    try
        let ret = cmd.ExecuteNonQuery() 
        let ms = stopWatch.Elapsed.TotalMilliseconds
        printfn "%.3f\t%s" ms qStr
        ms
    with 
        | ex // :? System.Exception as ex 
            ->  let errorMsg = ex.Message + "\r\n" + ex.StackTrace+ "\r\n" + qStr
                Console.WriteLine(errorMsg)
                0.0

// find the measure names from a simple MDX query. (It's probably easier to use AMO.)
let getMeasureNames () =
    let mdx = sprintf "SELECT {} on 0
	, ORDER(  [Measures].AllMembers
		    , [Measures].CurrentMember.MEMBER_NAME )  on 1
        FROM [%s] " cubeName 
    printfn "%s" mdx
    let cmd = adomdConn.CreateCommand()
    cmd.CommandText <- mdx
    cmd.CommandType <- CommandType.Text
    use rdr = cmd.ExecuteXmlReader()
    let xml = rdr.ReadOuterXml().ToString()
    let xdoc = new XmlDocument()
    let nsman = new XmlNamespaceManager(xdoc.NameTable)
    let xmlns="urn:schemas-microsoft-com:xml-analysis:mddataset" 
    nsman.AddNamespace("x", xmlns)
    xdoc.LoadXml(xml)
    let captionsXml = xdoc.SelectNodes("/x:root/x:Axes/x:Axis[@name='Axis1']/x:Tuples/x:Tuple/x:Member/x:Caption/text()", nsmgr = nsman)
    let captions = [| for n in captionsXml -> n.InnerText |]
    captions 

let measureNames = getMeasureNames ()

// this version timed the same measure 10 times in a row
//let timeMeasure m = 
//    let mdx = sprintf "SELECT [Measures].[%s] ON 0 FROM [%s]" m cubeName
//    clearCache ()
//    [| for i = 1 to 10 do yield timeMdx mdx |]
//let timings = 
//    [| for m in measureNames -> m,timeMeasure m |] 

// t1 = (snd timings).[0] = first run through all measures
// t2 = (snd timings).[1] = second run through all measures, etc
let timings = [| for m in getMeasureNames () do
                 yield m, (Array.zeroCreate numTimings  : double array) |]

let timeMeasure (mt: string * double array) mi = 
    let mdx = sprintf "SELECT [Measures].[%s] ON 0 FROM [%s]" (fst mt) cubeName
    (snd mt).[mi] <- timeMdx mdx

let fillTimings () = 
    for i = 0 to numTimings - 1 do
        printfn "Timing all measures, iteration #%d" i
        for mt in timings do timeMeasure mt i

fillTimings ()

//http://blogs.msdn.com/b/jackhu/
// read and write data to\from a excel workbook
#r "Microsoft.Office.Interop.Excel"
#r "office"
open Microsoft.Office.Interop

// Start Excel, Open a exiting file for input and create a new file for output
let xlApp = new Excel.ApplicationClass()
let xlWorkBook = xlApp.Workbooks.Add()
xlApp.Visible <- true
 
let tempDir = Environment.GetEnvironmentVariable("TEMP")

let outfile = sprintf @"%s\BenchmarkAllMeasures_%s_%s.xlsx" tempDir serverName asDbName

let mutable wfound = false
for w in xlApp.Workbooks do
    if w.FullName = outfile then
        w.Activate()
        wfound <- true

if not wfound then
    if File.Exists(outfile) then
        File.Delete(outfile)

// Open input's 'Sheet1' and create a new worksheet in output.xlsx
let xlws = xlWorkBook.Worksheets.[1] :?> Excel.Worksheet
//xlws.Name <- "OutputSheet1"

// fill in row headers
xlws.Cells.[1, 1] <- "MeasureName"
for c in 0 .. (snd timings.[0]).Length - 1 do
    xlws.Cells.[1, c + 2] <- sprintf "t%d" (c + 1)
// fill in data cells
for r in 0 .. timings.Length - 1 do
    let measureName,ts = timings.[r]
    xlws.Cells.[r + 2, 1] <- measureName
    for c in 0 .. ts.Length - 1 do
        xlws.Cells.[r + 2, c + 2] <- ts.[c]
    let addrPrevCell = xlws.Range("A1").Offset(r + 1, ts.Length).Address(false, false)
    //printfn "addrPrevCell=[%s]" addrPrevCell 
    xlws.Cells.[r + 2, ts.Length + 2] <- sprintf "=PERCENTILE.INC(B%d:%s,0.95)" (r + 2) addrPrevCell

xlws.UsedRange.Columns.NumberFormat <- "0"
xlws.UsedRange.Columns.AutoFit()
xlws.Range("A2").Activate()
xlApp.ActiveWindow.FreezePanes <- true
xlWorkBook.SaveAs(outfile) //, AccessMode = Excel.XlSaveAsAccessMode)

//xlWorkBook.Close() // I prefer to see my output
//xlApp.Quit()
