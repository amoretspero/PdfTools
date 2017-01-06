module ExtractPdf

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
    open System.Runtime.InteropServices

    let getFileSizeFormat (size : int64) =
        let sizes = [| "B"; "KB"; "MB"; "GB"; "TB" |]
        let mutable inputSize = (double)size
        let mutable order = 0
        while (inputSize >= 1024.0 && order < sizes.Length-1) do
            order <- order + 1
            inputSize <- inputSize / 1024.0
        String.Concat([| (inputSize.ToString(".00") :> obj); (sizes.[order] :> obj) |])


    let extractFormGenerator () =
        
        // Form definition region
        let extractForm = new MetroForm()
        extractForm.AutoSizeMode <- AutoSizeMode.GrowOnly
        extractForm.AutoSize <- true
        extractForm.Closing.Add(fun e ->
        if (extractForm.Opacity > 0.0) then
            e.Cancel <- true
            let extractFormClosingTimer = new Timer()
            extractFormClosingTimer.Interval <- 10
            extractFormClosingTimer.Tick.Add(fun _ ->
                if (extractForm.Opacity > 0.0) then extractForm.Opacity <- extractForm.Opacity - 0.1
                else 
                    extractFormClosingTimer.Stop()
                    extractForm.Close()
            )
            extractFormClosingTimer.Start()
        else
            extractForm.Close()
        )


        // File dialog definition region
        let openFileDialogAddPdf =
            new OpenFileDialog(
                InitialDirectory = "C:\\",
                Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                Multiselect = true
            )

        let saveFileDialog = 
            new SaveFileDialog(
                Filter = "pdf file|*.pdf",
                Title = "Save extracted file"
            )


        // Label definition region
        let label_OriginalPdf = new MetroLabel(Text = "Original PDF file", Width = 120, Height = 30, Location = new Point(50, 100), TextAlign = ContentAlignment.MiddleCenter)
        let label_OriginalPdfFileName = new MetroLabel(Text = "File name", Width = 100, Height = 30, Location = new Point(50, 150), TextAlign = ContentAlignment.MiddleCenter)
        let label_OriginalPdfFileLocation = new MetroLabel(Text = "File location", Width = 100, Height = 30, Location = new Point(50, 200), TextAlign = ContentAlignment.MiddleCenter)
        let label_OriginalPdfPageCount = new MetroLabel(Text = "Page count", Width = 100, Height = 30, Location = new Point(50, 250), TextAlign = ContentAlignment.MiddleCenter)
        let label_OriginalPdfFileSize = new MetroLabel(Text = "File size", Width = 100, Height = 30, Location = new Point(50, 300), TextAlign = ContentAlignment.MiddleCenter)

        let label_PdfFileName = new MetroLabel(Text = "", Width = 640, Height = 30, Location = new Point(200, 150), TextAlign = ContentAlignment.MiddleCenter)
        let label_PdfFileLocation = new MetroLabel(Text = "", Width = 640, Height = 30, Location = new Point(200, 200), TextAlign = ContentAlignment.MiddleCenter)
        let label_PdfPageCount = new MetroLabel(Text = "", Width = 640, Height = 30, Location = new Point(200, 250), TextAlign = ContentAlignment.MiddleCenter)
        let label_PdfFileSize = new MetroLabel(Text = "", Width = 640, Height = 30, Location = new Point(200, 300), TextAlign = ContentAlignment.MiddleCenter)

        let label_PagesToExtract = new MetroLabel(Text = "Pages to extract", Width = 120, Height = 30, Location = new Point(50, 400), TextAlign = ContentAlignment.MiddleCenter)

        extractForm.Controls
            .AddRange([| 
                        label_OriginalPdf; label_OriginalPdfFileName; label_OriginalPdfFileLocation; label_OriginalPdfPageCount; label_OriginalPdfFileSize; 
                        label_PagesToExtract;
                        label_PdfFileName; label_PdfFileLocation; label_PdfPageCount; label_PdfFileSize
            |])


        // Button definition region
        let button_SelectOriginalPdf = new MetroButton(Text = "Select PDF file", Width = 100, Height = 30, Location = new Point(200, 50))
        let button_ClearOriginalPdf = new MetroButton(Text = "Clear PDF file", Width = 100, Height = 30, Location = new Point(350, 50))
        let button_ExtractPdf = new MetroButton(Text = "Extract!", Width = 100, Height = 30, Location = new Point(50, 500))
        let button_EndProgram = new MetroButton(Text = "Close", Width = 70, Height = 30, Location = new Point(720, 500), Margin = new Padding(20))

        extractForm.Controls.AddRange([| button_SelectOriginalPdf; button_ClearOriginalPdf; button_ExtractPdf; button_EndProgram |])


        // Textbox definition region
        let textbox_OriginalPdf = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(200, 100), ReadOnly = true, Margin = new Padding(0, 0, 20, 0), AutoSize = true)
        let textbox_PagesToExtract = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(200, 400), TextAlign = HorizontalAlignment.Center, Margin = new Padding(0, 0, 20, 0), AutoSize = true)

        extractForm.Controls.AddRange([| textbox_OriginalPdf; textbox_PagesToExtract |])


        // Button event listener region.
        button_SelectOriginalPdf.Click.Add(fun _ ->
            let openFileDialogAddPdfResult = openFileDialogAddPdf.ShowDialog()
            match openFileDialogAddPdfResult with
            | DialogResult.OK ->
                let fileFullName = openFileDialogAddPdf.FileName
                let fileName = Path.GetFileName(openFileDialogAddPdf.FileName)
                let fileLocation = Path.GetDirectoryName(openFileDialogAddPdf.FileName)
                let fileSize = FileInfo(openFileDialogAddPdf.FileName).Length |> getFileSizeFormat
                if (PdfReader.TestPdfFile(fileFullName) = 0) then
                    MetroMessageBox.Show(extractForm, "Input PDF file is invalid", "Error - Invalid PDF file(s)", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                else
                    let pdfDoc = PdfReader.Open(fileFullName, PdfDocumentOpenMode.Import)
                    let pageCount = pdfDoc.Pages.Count
                    label_PdfFileName.Text <- fileName
                    label_PdfFileLocation.Text <- fileLocation
                    label_PdfFileSize.Text <- fileSize
                    label_PdfPageCount.Text <- pageCount.ToString()
                    textbox_OriginalPdf.Text <- fileFullName
            | _ ->
                printfn "%A" openFileDialogAddPdfResult
        )

        button_ClearOriginalPdf.Click.Add(fun _ ->
            label_PdfFileName.Text <- ""
            label_PdfFileLocation.Text <- ""
            label_PdfFileSize.Text <- ""
            label_PdfPageCount.Text <- ""
            textbox_OriginalPdf.Text <- ""
        )

        button_EndProgram.Click.Add(fun _ ->
            extractForm.Close()
        )

        button_ExtractPdf.Click.Add(fun _ ->
            let fileFullName = textbox_OriginalPdf.Text
            if (not (File.Exists(fileFullName))) then
                MetroMessageBox.Show(extractForm, "Input PDF file does not exists.", "Error - Invalid input file", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (PdfReader.TestPdfFile(fileFullName) = 0) then
                MetroMessageBox.Show(extractForm, "Input PDF file is invalid.", "Error - Invalid input file", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let saveFileDialogResult = saveFileDialog.ShowDialog()
                match saveFileDialogResult with
                | DialogResult.OK ->
                    let extractedPdfFileName = saveFileDialog.FileName
                    let extractedPdfDoc = new PdfDocument()
                    let originalPdfDoc = PdfReader.Open(fileFullName, PdfDocumentOpenMode.Import)
                    let pages = textbox_PagesToExtract.Text.Split([| ","; " " |], StringSplitOptions.RemoveEmptyEntries)
                    let mutable isPagesValid = true
                    for pg in pages do
                        try
                            let converted = System.Convert.ToInt32(pg)
                            if (converted > originalPdfDoc.Pages.Count || converted <= 0) then isPagesValid <- false
                            ()
                        with
                            | :? System.Exception as e -> 
                                isPagesValid <- false
                    if (not isPagesValid) then MetroMessageBox.Show(extractForm, "Input pages are not correct.", "Error - Invalid input page", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                    for pg in pages do
                        extractedPdfDoc.Pages.Add(originalPdfDoc.Pages.[(System.Convert.ToInt32(pg))-1]) |> ignore
                    extractedPdfDoc.Save(extractedPdfFileName)
                    MetroMessageBox.Show(extractForm, "Successfully extracted PDF file.", "Success - Extract", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore
                | _ ->
                    printfn "%A" saveFileDialogResult
        )

        extractForm