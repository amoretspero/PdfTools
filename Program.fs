// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.
open MergePdf
open MergePdfMultiple
open ExtractPdf
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

    let mergeTile = new MetroTile(Text = "Merge PDF\n - 2 Files", Size = new Size(120, 120), TextAlign = ContentAlignment.MiddleCenter, Location = new Point(120, 120))
    mergeTile.Click.Add(fun _ ->
        mainForm.Hide()
        let mergeForm = mergeFormGenerator()
        mergeForm.FormClosed.Add(fun _ ->
            mainForm.Show()
        )
        mergeForm.ShowDialog() |> ignore
    )

    let mergeMultipleTile = new MetroTile(Text = "Merge PDF\n - Multiple Files", Size = new Size(120, 120), TextAlign = ContentAlignment.MiddleCenter, Location = new Point(260, 120))
    mergeMultipleTile.Click.Add(fun _ ->
        mainForm.Hide()
        let mergeMultipleForm = mergeMultipleFormGenerator()
        mergeMultipleForm.FormClosed.Add(fun _ ->
            mainForm.Show()
        )
        mergeMultipleForm.ShowDialog() |> ignore
    )

    let extractTile = new MetroTile(Text = "Extract PDF", Size = new Size(120, 120), TextAlign = ContentAlignment.MiddleCenter, Location = new Point(400, 120))
    extractTile.Click.Add(fun _ ->
        mainForm.Hide()
        let extractForm = extractFormGenerator()
        extractForm.FormClosed.Add(fun _ ->
            mainForm.Show()
        )
        extractForm.ShowDialog() |> ignore
    )
    
    mainForm.Closing.Add(fun e ->
        if (mainForm.Opacity > 0.0) then
            e.Cancel <- true
            let mainFormClosingTimer = new Timer()
            mainFormClosingTimer.Interval <- 10
            mainFormClosingTimer.Tick.Add(fun _ ->
                if (mainForm.Opacity > 0.0) then mainForm.Opacity <- mainForm.Opacity - 0.1
                else 
                    mainFormClosingTimer.Stop()
                    mainForm.Close()
            )
            mainFormClosingTimer.Start()
        else
            mainForm.Close()
    )
    
    
    mainForm.Controls.AddRange([| mergeTile; mergeMultipleTile; extractTile |])

    mainForm.ShowDialog() |> ignore

    0 // return an integer exit code
