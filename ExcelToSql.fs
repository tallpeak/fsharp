// This can run as an fsx script; see ExcelToSQL.bat (no EXE required, but fsi.exe and FSharp.Core.dll must exist somewhere)

// Generate a SQL script from a worksheet of an Excel file
// by Aaron W. West, 5/31/2014

// Usage: excelToSQL databaseName schemaName tableName infile.xls outfile.sql
// Name a sheet of the workbook *[tableName]* or *data* or Sheet1,
// Make sure your ActiveRange has no blank columns (delete blank columns or rows if so),
// Make sure the formatting of the cells is appropriate for the datatype
// Make sure the first row contains field names for the table   

// This uses:
// https://code.google.com/p/linqtoexcel/
// http://www.nuget.org/packages/LinqToExcel  ( Install-Package LinqToExcel )
#if INTERACTIVE
#I "..\packages\LinqToExcel.1.8.0\lib" //assembly search path
#r "LinqToExcel.dll"
#r "Remotion.Data.Linq.dll"
#endif

open System
open LinqToExcel
open LinqToExcel.Query
open LinqToExcel.Extensions
open LinqToExcel.Attributes
open LinqToExcel.Domain

let excelToSql databaseName schemaName tableName fn outfn =
    //let ci = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
    //let dtfi: System.Globalization.DateTimeFormatInfo = ci.DateTimeFormat
    let fullTableName = sprintf "[%s].[%s].[%s]" databaseName schemaName tableName
    let excel = new ExcelQueryFactory(fn)   
    excel.DatabaseEngine <- DatabaseEngine.Ace
    let isDataSheet n = let nm = n.ToString().ToLower() in 
                        nm.Contains(tableName) || nm.Contains("data") || nm.Equals("sheet1")
    let firstWorksheetName = excel.GetWorksheetNames() |> Seq.where(isDataSheet) |> Seq.take(1) |> Seq.exactlyOne 
    let q = query { for row in excel.Worksheet(firstWorksheetName) do
                    select row }
    let colNames : string[] = excel.GetColumnNames (firstWorksheetName) |> Seq.toArray //<string>
    let colNamesList = colNames |> Array.toList
    // create INSERT 
    let insertStart = sprintf "INSERT %s (%s)\n" fullTableName (String.Join(",", colNames ))
    // Terminate if we find an unknown type, so the programmer can fix the problem
    // However, LinqToExcel is returning a very limited set of types, not all worksheet cell types
    let formatCell (v:obj) = 
        match v with 
        | :? System.String   -> sprintf "'%s'" <| v.ToString().Replace("'","''") 
        | :? System.DateTime -> sprintf "'%s'" <| v.ToString().Replace("'","''") 
        | :? System.Double   -> sprintf "%s" <| v.ToString()
        | :? System.DBNull   -> "NULL"  
//      | :? System.Int32    -> sprintf "%s" <| v.ToString() // Spreadsheets don't have Int32
        | _ -> raise(Exception("Error: Type=" + v.GetType().ToString() 
                                + "???: Unknown type in formatCell")) 
    let formatRow (r:Row) = 
        sprintf "%s VALUES(%s)" insertStart
        <| String.Join(",", [| for cel in r.ToArray()  
                               -> formatCell (cel.Value) |])
 
    let inserts = q |> Seq.map( formatRow ) |> Seq.toArray
    
    let lengthCell (v:obj) = 
        match v with 
        | :? System.String   -> v.ToString().Length
        | :? System.DateTime -> 8 
        | :? System.Double   -> v.ToString().Length
        | :? System.DBNull   -> 0  
        | _ -> raise(Exception("Error: Type=" + v.GetType().ToString() 
                                + "???: Unknown type in lengthCell")) 
  
    // Convert a list of lists of column lengths into a list of maximum column lengths
    // incomplete pattern-match if the lists are uneven 
    // (this assumes rows of equal-length lists of cells)
    let rec maxLists (ill : list<list<int>>) = 
        match List.head ill with 
            | hd::tl -> List.Cons ( List.max <| List.map List.head ill, 
                                    maxLists (List.map List.tail ill) )
            | []     -> []

    let cellLengths (r:Row) = [ for cel in r.ToArray() -> lengthCell (cel.Value) ]
    let qlens = q |> Seq.toList |> List.map( cellLengths ) 
    let maxLengths = maxLists qlens
    let maxLengthStrings = List.map (sprintf "%d") maxLengths
    let zipWith op x y = List.map (fun (x,y) -> op x y) <| List.zip x y 
    let formatter colname len = sprintf "%s\t%s" colname len
    let maxLengthsString = String.Join("\n", zipWith formatter colNamesList maxLengthStrings)
    let prefix = String.Format("USE {0}\nGO\nSET IDENTITY_INSERT {1}.{2} ON\nGO\n",
                                databaseName, schemaName, tableName)
    let output = sprintf "%s%s\n/* Maximum cell lengths:\n%s\n*/" 
                          prefix (String.Join("\n",inserts)) maxLengthsString
    // Warning, side-effects below (not so functional)

    if outfn = "-" then 
        stdout.Write(output)
    else
        let outfile = System.IO.File.WriteAllText(outfn, output)
        printfn "Opening %s ..." outfn

        // you might want to remove this statement, 
        // but I think it's cool to open the resulting sql file in SSMS 
        // so it can be run or tested immediately
        System.Diagnostics.Process.Start(outfn) |> ignore

// Replace extension, eg filename.xlsx to filename.sql. 
// Arg ext should contain the dot
let replaceExt (fn:string) (ext:string) = fn.Split([|'.'|]).[0] + ext

let doMain argv = 
    match argv with 
        | [|databaseName; schemaName; tableName; fn; outfn|] 
            -> excelToSql databaseName schemaName tableName fn outfn
        | [|databaseName; schemaName; tableName; fn|] 
            -> excelToSql databaseName schemaName tableName fn (replaceExt fn ".sql")
        | [|"DEBUGTITAN"|] 
            -> excelToSql "IDODS_ODS" "dbo" "ODS_Trade" 
                           @"C:\titan\data\IDODS_TradeLoad_ODS_Trades.xls" 
                           @"C:\titan\data\IDODS_TradeLoad_ODS_Trades.sql"
        | _ -> printfn "Usage: excelToSQL databaseName schemaName tableName infile.xls outfile.sql\n%s"
                        ", or substitute - for outfile.sql for stdout"
               exit(1)
    0 // return an integer exit code


#if COMPILED
[<EntryPoint>]
let main argv = 
    doMain argv
#else

// Three ways to run the program
// Exe: simplest, actually. But it does require Fsharp.core.dll to be in the path, I think

// alternative active patterns syntax
//let (|FileExtension|) (filePath:string) = IO.Path.GetExtension(filePath).ToLower()
//let (|IsScriptName|) (filePath:string) = filePath.ToLower().EndsWith(".fsx")

//from http://www2.lib.uchicago.edu/keith/ocaml-class/utils.html
let rec dropwhile f = function
  | [] -> []
  | hd::tl when f hd -> dropwhile f tl
  | list -> list

let dropWhileStrings = (dropwhile : (string -> bool) -> string list -> string list)

// GetCommandLineArgs, then drop *fsi.exe (optionally), then drop *.fsx
let args = dropWhileStrings 
            (fun s -> not ( s.ToLower().EndsWith(".fsx")) ) 
            // with active patterns it could be
            // (fun s -> match s with | FileExtension ".fsx" -> false | _ -> true)
            (Array.toList( Environment.GetCommandLineArgs() ))
            |> List.tail |> List.toArray
printfn "args=%A" args
doMain args  
// or // excelToSql (args.[0]) (args.[1]) (args.[2]) (args.[3]) (args.[4]) 
#endif

