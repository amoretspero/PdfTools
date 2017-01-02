module MergePdf

    open System
    open System.Collections
    open System.Collections.Generic
    open System.Diagnostics
    open System.IO
    open System.Linq
    open System.Text
    open System.Threading
    open System.Drawing
    open System.Windows.Forms
    open System.Windows.Forms.DataVisualization
    open PdfSharp
    open PdfSharp.Pdf
    open PdfSharp.Pdf.IO
    open MetroFramework
    open MetroFramework.Controls
    open MetroFramework.Drawing
    open MetroFramework.Forms

    let mergeFormGenerator () =

        // Form definition region.
    
        let mergeForm = new MetroForm()
        mergeForm.AutoSizeMode <- AutoSizeMode.GrowOnly
        mergeForm.AutoSize <- true

        // File dialog definition region.
        let openFileDialogPdf1 = 
            new OpenFileDialog(
                InitialDirectory = "C:\\", 
                Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            )

        let openFileDialogPdf2 = 
            new OpenFileDialog(
                InitialDirectory = "C:\\",
                Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            )

        let saveFileDialog = 
            new SaveFileDialog(
                Filter = "pdf file|*.pdf",
                Title = "Save merged file"
            )
    
        // Label definition region.
        let label_InputPdf1 = new MetroLabel(Text = "Input 1", Width = 100, Height = 30, Location = new Point(20, 50), TextAlign = ContentAlignment.MiddleCenter)
        let label_InputPdf2 = new MetroLabel(Text = "Input 2", Width = 100, Height = 30, Location = new Point(20, 200), TextAlign = ContentAlignment.MiddleCenter)

        let label_Status = new MetroLabel(Text = "Ready", Width = 100, Height = 30, Location = new Point(20, 380), TextAlign = ContentAlignment.MiddleCenter)

        mergeForm.Controls.AddRange([| label_InputPdf1; label_InputPdf2; label_Status |])


        // Button definition region.
        let button_InputPdf1 = new MetroButton(Text = "Select Pdf 1", Width = 100, Height = 30, Location = new Point(170, 50))
        let button_InputPdf2 = new MetroButton(Text = "Select Pdf 2", Width = 100, Height = 30, Location = new Point(170, 200))
    
        let button_ResetPdf1 = new MetroButton(Text = "Reset Pdf 1", Width = 100, Height = 30, Location = new Point(320, 50))
        let button_ResetPdf2 = new MetroButton(Text = "Reset Pdf 2", Width = 100, Height = 30, Location = new Point(320, 200))

        let button_MergePdf = new MetroButton(Text = "Merge!", Width = 100, Height = 30, Location = new Point(20, 320))

        let button_EndProgram = new MetroButton(Text = "Close", Width = 70, Height = 30, Location = new Point(760, 420), Margin = new Padding(20))

        mergeForm.Controls.AddRange([| button_InputPdf1; button_InputPdf2; button_ResetPdf1; button_ResetPdf2; button_EndProgram; button_MergePdf |])


        // Textbox definition region.
        let textbox_InputPdf1 = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(170, 100), ReadOnly = true, Margin = new Padding(0, 0, 20, 0), AutoSize = true)
        let textbox_InputPdf2 = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(170, 250), ReadOnly = true, Margin = new Padding(0, 0, 20, 0), AutoSize = true)

        mergeForm.Controls.AddRange([| textbox_InputPdf1; textbox_InputPdf2 |])

        // Button event listener region.
        button_InputPdf1.Click.Add(fun _ ->
            let openFileDialogPdf1Result = openFileDialogPdf1.ShowDialog()
            match openFileDialogPdf1Result with
            | DialogResult.OK ->
                try
                    let fileName = openFileDialogPdf1.FileName
                    textbox_InputPdf1.Text <- fileName
                with
                    | :? System.Exception as e -> printfn "Error: %s" e.Message
            | DialogResult.Cancel ->
                printfn "Cancelled selection."
            | _ ->
                printfn "%A" openFileDialogPdf1Result
        )

        button_ResetPdf1.Click.Add(fun _ ->
            textbox_InputPdf1.Text <- ""
        )

        button_InputPdf2.Click.Add(fun _ ->
            let openFileDialogPdf2Result = openFileDialogPdf2.ShowDialog()
            match openFileDialogPdf2Result with
            | DialogResult.OK ->
                try
                    let fileName = openFileDialogPdf2.FileName
                    textbox_InputPdf2.Text <- fileName
                with
                    | :? System.Exception as e -> printfn "Error: %s" e.Message
            | DialogResult.Cancel ->
                printfn "Cancelled selection."
            | _ ->
                printfn "%A" openFileDialogPdf2Result
        )

        button_ResetPdf2.Click.Add(fun _ ->
            textbox_InputPdf2.Text <- ""
        )

        button_EndProgram.Click.Add(fun _ ->
            mergeForm.Close()
        )

        button_MergePdf.Click.Add(fun _ ->
            if (textbox_InputPdf1.Text = "") then
                MetroMessageBox.Show(mergeForm, "Please select input source 1.", "Error - Unspecified source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (textbox_InputPdf2.Text = "") then
                MetroMessageBox.Show(mergeForm, "Please select input source 2.", "Error - Unspecified source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (PdfReader.TestPdfFile(textbox_InputPdf1.Text) = 0) then
                MetroMessageBox.Show(mergeForm, "Source 1 is invalid pdf file.", "Error - Invalid source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (PdfReader.TestPdfFile(textbox_InputPdf2.Text) = 0) then
                MetroMessageBox.Show(mergeForm, "Source 2 is invalid pdf file.", "Error - Invalid source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let saveFileDialogResult = saveFileDialog.ShowDialog()
                match saveFileDialogResult with
                | DialogResult.OK ->
                    label_Status.Text <- "Reading Source 1..."
                    let pdfSource1 = PdfReader.Open(textbox_InputPdf1.Text, PdfDocumentOpenMode.Import)
                    label_Status.Text <- "Reading Source 2..."
                    let pdfSource2 = PdfReader.Open(textbox_InputPdf2.Text, PdfDocumentOpenMode.Import)
                    printfn "Successfully opend pdf file sources."
                    let newName = saveFileDialog.FileName
                    label_Status.Text <- "Merging files..."
                    let pdfMerged = new PdfDocument()
                    for i=0 to pdfSource1.PageCount-1 do
                        pdfMerged.AddPage(pdfSource1.Pages.[i]) |> ignore
                    for i=0 to pdfSource2.PageCount-1 do
                        pdfMerged.AddPage(pdfSource2.Pages.[i]) |> ignore
                    pdfMerged.Save(newName)
                    printfn "Successfully saved merged pdf file."
                    label_Status.Text <- "Success!"
                    MetroMessageBox.Show(mergeForm, "Successfully merged files.", "Success - Merge", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore
                    label_Status.Text <- "Ready"
                | DialogResult.Cancel ->
                    printfn "Cancelled merge."
                | _ ->
                    printfn "%A" saveFileDialogResult
        )

        mergeForm