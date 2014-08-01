// Generate a SQL script from the first worksheet of an Excel file
// by Aaron W. West 
// Usage: excelToSQL databaseName schemaName tableName infile.xls outfile.sql

// This assumes the first sheet of the workbook contains a table with no blank columns
// and the first column is the field names of the table   

// You may need to convert some columns to all-strings for LinqToExcel to infer the correct type for the column

// TODO: ??  check schema of destination table so that 
// formatCell can use the type of the destination column
// to determine how to format properly , and output an error to the user
// if the SQL would be invalid (too long for destination column), 
// and can avoid inserting NULL into NOT NULL columns

// TODO: generate a suggested schema when no schema is given. 
// TODO? optional Values, ...  format
// TODO? Install-Package UnionArgParser

// This uses:
// https://code.google.com/p/linqtoexcel/
// http://www.nuget.org/packages/LinqToExcel  ( Install-Package LinqToExcel )

#I "..\packages\LinqToExcel.1.8.0\lib" //assembly search path
#r "LinqToExcel.dll"
#r "Remotion.Data.Linq.dll"
open System
open LinqToExcel
open LinqToExcel.Query
open LinqToExcel.Extensions
open LinqToExcel.Attributes
open LinqToExcel.Domain

let excelToSql databaseName schemaName tableName fn outfn =
    let ci = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
    let dtfi: System.Globalization.DateTimeFormatInfo = ci.DateTimeFormat
    let fullTableName = sprintf "[%s].[%s].[%s]" databaseName schemaName tableName
    let excel = new ExcelQueryFactory(fn)   
    excel.DatabaseEngine <- DatabaseEngine.Ace
    let worksheetNames = query { for wn in excel.GetWorksheetNames() do select wn.ToString }
    let firstWorksheetName = worksheetNames |> Seq.take(1) |> Seq.exactlyOne  <| dtfi
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
//      | :? System.Int32    -> sprintf "%s" <| v.ToString()
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
//    Console.WriteLine("press enter to exit")
//    Console.Read() |> ignore
    //0 // return an integer exit code


#if COMPILED
[<EntryPoint>]
let main argv = 
    doMain argv
#else // #if INTERACTIVE

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
#endif

