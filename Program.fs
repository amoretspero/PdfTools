// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.
open MergePdf
open MetroFramework
open MetroFramework.Controls
open MetroFramework.Forms
open System
open System.Threading
open System.Drawing
open System.Windows.Forms
open System.Windows.Forms.DataVisualization
open System.Runtime.InteropServices

[<DllImport("user32.dll")>]
extern bool SetProcessDPIAware();

[<EntryPoint>]
[<STAThreadAttribute>]
let main argv = 
    printfn "%A" argv

    //if (System.Environment.OSVersion.Version.Major >= 6) then
    //    SetProcessDPIAware() |> ignore
 
    //Application.EnableVisualStyles()
    //Application.SetCompatibleTextRenderingDefault(true)

    let mainForm = new MetroForm(Text = "PdfTools - Simple PDF management tool.")
    mainForm.AutoSizeMode <- AutoSizeMode.GrowOnly
    mainForm.AutoSize <- true
    mainForm.MinimumSize <- new Size(640, 480)

    let mergeTile = new MetroTile(Text = "Merge PDF", Size = new Size(120, 120), TextAlign = ContentAlignment.MiddleCenter, Location = new Point(120, 120))
    mergeTile.Click.Add(fun _ ->
        mainForm.Hide()
        let mergeForm = mergeFormGenerator()
        mergeForm.FormClosed.Add(fun _ ->
            
            mainForm.Show()
        )
        mergeForm.ShowDialog() |> ignore
    )

    mainForm.Controls.AddRange([| mergeTile |])

    mainForm.ShowDialog() |> ignore

    0 // return an integer exit code
