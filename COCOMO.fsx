// inspired by looking at an open-source project and wondering how its cost was calculated
// specifically https://www.openhub.net/p/bigdata
// I didn't quite get my numbers to match (off by about 10%)
// See https://en.wikipedia.org/wiki/COCOMO#Basic_COCOMO
#r "PresentationCore.dll"
#r "PresentationFramework.dll"
#r "System.dll"
#r "System.Core.dll"
#r "System.Numerics.dll"
#r "System.Xaml.dll"
#r "System.Xml.dll"
#r "WindowsBase.dll"

open System
open System.Windows
open System.Windows.Controls
open System.Xml
open System.Windows.Markup

let xaml = """<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:fsxaml="clr-namespace:FsXaml;assembly=FsXaml.Wpf"
        xmlns:local="assembly=ConsoleApplication1"

        Title="MainWindow" Height="350" Width="525">
    <Grid>
        <TextBox x:Name="txtLOC" HorizontalAlignment="Left" Height="23" Margin="82,43,0,0" TextWrapping="Wrap" Text="0" VerticalAlignment="Top" Width="120"/>
        <TextBox x:Name="txtCostPerManYear" HorizontalAlignment="Left" Height="23" Margin="82,81,0,0" TextWrapping="Wrap" Text="55000" VerticalAlignment="Top" Width="120"/>
        <Button x:Name="buttonRecalc" Content="Recalc" HorizontalAlignment="Left" Margin="78,238,0,0" VerticalAlignment="Top" Width="75" />
        <Label x:Name="label" Content="lines of code" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="225,41,0,0"/>
        <Label x:Name="labelDollarsPerManYear" Content="$/man year" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="225,82,0,0"/>
        <Label x:Name="outManYears" Content="ManYears" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="3.306,4.856" Margin="82,121,0,0" Width="120"/>        
        <Label x:Name="outCost" Content="Cost" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="3.306,4.856" Margin="82,158,0,0" Width="120"/>
        <Label x:Name="labelManYearsLabel" Content="ManYears" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="3.306,4.856" Margin="225,121,0,0" Width="120"/>
        <Label x:Name="label1_Copy3" Content="Cost" HorizontalAlignment="Left" VerticalAlignment="Top" RenderTransformOrigin="3.306,4.856" Margin="225,158,0,0" Width="120"/>

    </Grid>
</Window>
"""

//http://fssnip.net/H/title/Load-XAML
let loadXamlWindow (xaml:string) =
    let strm = new IO.StringReader(xaml)
    let reader = XmlReader.Create(strm)
    let w = System.Windows.Markup.XamlReader.Load(reader) :?> Window
    w

let buttonClick  (w : Window) (e : RoutedEventArgs ) = 
    let txtLOC = w.FindName("txtLOC") :?> TextBox
    let txtCostPerManYear = w.FindName("txtCostPerManYear") :?> TextBox
    let outManYears = w.FindName("outManYears") :?> Label
    let outCost = w.FindName("outCost") :?> Label
    let loc = match Double.TryParse(txtLOC.Text) with | true,v -> v | _ -> 0.0 
    let kloc = loc * 0.001
    let costPerManYear = match Double.TryParse(txtCostPerManYear.Text) with | true, v -> v | _ -> 0.0
    //Organic // https://en.wikipedia.org/wiki/COCOMO#Basic_COCOMO
    let ab = 2.4 // 2.325?
    let bb = 1.05
    let manMonths = ab * Math.Pow(kloc, bb)
    let manYears = manMonths / 12.0 
    let cost = manYears * costPerManYear
    outManYears.Content <- manYears
    outCost.Content <- cost

let private main (args: string []) =
    let mainWindow = loadXamlWindow(xaml)
    let app= Application()
    let b = mainWindow.FindName("buttonRecalc") :?> Button
    b.Click.Add (buttonClick mainWindow)
    app.Run(mainWindow)

#if INTERACTIVE
fsi.CommandLineArgs |> Array.toList |> List.tail |> List.toArray |> main
#else
[<EntryPoint; STAThread>]
let entryPoint args = main args
#endif