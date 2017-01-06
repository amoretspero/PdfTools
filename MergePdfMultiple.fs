module MergePdfMultiple

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

    let mergeMultipleFormGenerator () =

        // Form definition region.
    
        let mergeMultipleForm = new MetroForm()
        mergeMultipleForm.AutoSizeMode <- AutoSizeMode.GrowOnly
        mergeMultipleForm.AutoSize <- true
        mergeMultipleForm.Closing.Add(fun e ->
        if (mergeMultipleForm.Opacity > 0.0) then
            e.Cancel <- true
            let mergeMultipleFormClosingTimer = new Timer()
            mergeMultipleFormClosingTimer.Interval <- 10
            mergeMultipleFormClosingTimer.Tick.Add(fun _ ->
                if (mergeMultipleForm.Opacity > 0.0) then mergeMultipleForm.Opacity <- mergeMultipleForm.Opacity - 0.1
                else 
                    mergeMultipleFormClosingTimer.Stop()
                    mergeMultipleForm.Close()
            )
            mergeMultipleFormClosingTimer.Start()
        else
            mergeMultipleForm.Close()
        )

        // File dialog definition region.
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
                Title = "Save merged file"
            )
    
        // Label definition region.
        let label_InputPdf1 = new MetroLabel(Text = "Input 1", Width = 100, Height = 30, Location = new Point(20, 50), TextAlign = ContentAlignment.MiddleCenter)
        let label_InputPdf2 = new MetroLabel(Text = "Input 2", Width = 100, Height = 30, Location = new Point(20, 200), TextAlign = ContentAlignment.MiddleCenter)
        let label_InputPdfs = new MetroLabel(Text = "Input Pdf Files", Width = 100, Height = 30, Location = new Point(20, 50), TextAlign = ContentAlignment.MiddleCenter)

        let label_Status = new MetroLabel(Text = "Ready", Width = 100, Height = 30, Location = new Point(20, 200), TextAlign = ContentAlignment.MiddleCenter)

        mergeMultipleForm.Controls.AddRange([| label_InputPdfs; label_Status |])


        // Button definition region.
        let button_InputPdf1 = new MetroButton(Text = "Select Pdf 1", Width = 100, Height = 30, Location = new Point(170, 50))
        let button_InputPdf2 = new MetroButton(Text = "Select Pdf 2", Width = 100, Height = 30, Location = new Point(170, 200))
    
        let button_ResetPdf1 = new MetroButton(Text = "Reset Pdf 1", Width = 100, Height = 30, Location = new Point(320, 50))
        let button_ResetPdf2 = new MetroButton(Text = "Reset Pdf 2", Width = 100, Height = 30, Location = new Point(320, 200))

        let button_ClearSelected = new MetroButton(Text = "Clear selected item", Width = 120, Height = 30, Location = new Point(300, 50))
        let button_ClearAll = new MetroButton(Text = "Clear All items", Width = 100, Height = 30, Location = new Point(450, 50))

        let button_AddPdf = new MetroButton(Text = "Add Pdf file", Width = 100, Height = 30, Location = new Point(150, 50))

        let button_MergePdf = new MetroButton(Text = "Merge!", Width = 100, Height = 30, Location = new Point(650, 50))

        let button_EndProgram = new MetroButton(Text = "Close", Width = 70, Height = 30, Location = new Point(760, 420), Margin = new Padding(20))

        mergeMultipleForm.Controls.AddRange([| (*button_InputPdf1; button_InputPdf2; button_ResetPdf1; button_ResetPdf2;*) button_AddPdf; button_ClearSelected; button_ClearAll; button_EndProgram; button_MergePdf |])


        // Textbox definition region.
        let textbox_InputPdf1 = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(170, 100), ReadOnly = true, Margin = new Padding(0, 0, 20, 0), AutoSize = true)
        let textbox_InputPdf2 = new MetroTextBox(Text = "", MinimumSize = new Size(720, 30), Location = new Point(170, 250), ReadOnly = true, Margin = new Padding(0, 0, 20, 0), AutoSize = true)

        //mergeForm.Controls.AddRange([| textbox_InputPdf1; textbox_InputPdf2 |])

        // Listview definition region.
        let listview_Input = new ListView(Location = new Point(150, 100), Size = new Size(640, 120))
        listview_Input.View <- View.Details
        listview_Input.GridLines <- true
        listview_Input.FullRowSelect <- true
        listview_Input.Columns.Add("FileName", 300, HorizontalAlignment.Center) |> ignore
        listview_Input.Columns.Add("Location", 300, HorizontalAlignment.Center) |> ignore
        listview_Input.AllowDrop <- true
        listview_Input.ItemDrag.Add(fun _ ->
            listview_Input.DoDragDrop(listview_Input.SelectedItems, DragDropEffects.Move) |> ignore
        )
        listview_Input.DragDrop.Add(fun e ->
            //e.Effect <- DragDropEffects.Copy
            if (listview_Input.SelectedItems.Count > 0) then
                let cp = listview_Input.PointToClient(new Point(e.X, e.Y))
                let dragToItem = listview_Input.GetItemAt(cp.X, cp.Y)
                if (dragToItem <> null) then
                    let dragIndex = dragToItem.Index
                    let sel = Array.init listview_Input.SelectedItems.Count (fun x -> new ListViewItem())
                    for i=0 to listview_Input.SelectedItems.Count-1 do
                        sel.[i] <- listview_Input.SelectedItems.[i]
                    for i=0 to sel.GetLength(0) - 1 do
                        let dragItem = sel.[i]
                        let mutable itemIndex = dragIndex
                        if (itemIndex = dragItem.Index) then ()
                        if (dragItem.Index < itemIndex) then itemIndex <- itemIndex + 1
                        else itemIndex <- dragIndex + i
                        let insertItem = dragItem.Clone() :?> ListViewItem
                        listview_Input.Items.Insert(itemIndex, insertItem) |> ignore
                        listview_Input.Items.Remove(dragItem)
        )
        listview_Input.DragEnter.Add(fun e ->
            //listview_Input.Items.Add(e.Data.ToString()) |> ignore
            let len = e.Data.GetFormats().Length - 1
            for i=0 to len do
                if (e.Data.GetFormats().[i].Equals("System.Windows.Forms.ListView+SelectedListViewItemCollection")) then
                    e.Effect <- DragDropEffects.Move
        )

        mergeMultipleForm.Controls.AddRange([| listview_Input |])

        // Button event listener region.
        (*button_InputPdf1.Click.Add(fun _ ->
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
        )*)

        button_AddPdf.Click.Add(fun _ ->
            let openFileDialogAddPdfResult = openFileDialogAddPdf.ShowDialog()
            match openFileDialogAddPdfResult with
            | DialogResult.OK ->
                try
                    let files = openFileDialogAddPdf.FileNames
                    listview_Input.BeginUpdate()
                    for f in files do
                        let fileName = Path.GetFileName(f)
                        let location = Path.GetDirectoryName(f)
                        listview_Input.Items.Add(new ListViewItem([| fileName; location |])) |> ignore
                    listview_Input.EndUpdate()
                    //let fileName = Path.GetFileName(openFileDialogAddPdf.FileName)
                    //let location = Path.GetDirectoryName(openFileDialogAddPdf.FileName)
                    //listview_Input.Items.Add(new ListViewItem([| fileName; location |])) |> ignore
                    //listview_Input.UpdateScrollbar()
                with
                    | :? System.Exception as e -> printfn "Error: %s" e.Message
            | DialogResult.Cancel ->
                printfn "Cancelled selection."
            | _ ->
                printfn "%A" openFileDialogAddPdf
        )

        button_ClearSelected.Click.Add(fun _ ->
            let selectedItems = listview_Input.SelectedItems
            listview_Input.BeginUpdate()
            for si in selectedItems do
                listview_Input.Items.Remove(si)
            listview_Input.EndUpdate()
        )

        button_ClearAll.Click.Add(fun _ ->
            let items = listview_Input.Items
            listview_Input.BeginUpdate()
            for i in items do
                listview_Input.Items.Remove(i)
            listview_Input.EndUpdate()
        )

        button_EndProgram.Click.Add(fun _ ->
            mergeMultipleForm.Close()
        )

        button_MergePdf.Click.Add(fun _ ->
            if (listview_Input.Items.Count <= 1) then
                MetroMessageBox.Show(mergeMultipleForm, "Please add more than 1 file(s) to merge.", "Error - Not enough files", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else
                let invalidFiles = ref ([| |] : string [])
                for li in listview_Input.Items do
                    let fileName = li.SubItems.[0].Text
                    let path = li.SubItems.[1].Text
                    let fullPath = Path.Combine(path, fileName)
                    if (PdfReader.TestPdfFile(fullPath) = 0) then
                        invalidFiles.Value <- Array.append invalidFiles.Value [| fileName |]
                if (invalidFiles.Value.Count() > 0) then
                    let errorMessage = 
                        let sb = StringBuilder()
                        sb.Append("There are invalid pdf files: \n") |> ignore
                        for invalidFile in invalidFiles.Value do
                            sb.Append(invalidFile + "\n") |> ignore
                        sb.ToString()
                    MetroMessageBox.Show(mergeMultipleForm, errorMessage, "Error - Invalid PDF file(s)", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
                else
                    let saveFileDialogResult = saveFileDialog.ShowDialog()
                    match saveFileDialogResult with
                    | DialogResult.OK ->
                        label_Status.Text <- "Reading Sources..."
                        label_Status.Refresh() 
                        let pdfSources = ref ([| |] : PdfDocument [])
                        for li in listview_Input.Items do
                            let fileName = li.SubItems.[0].Text
                            let path = li.SubItems.[1].Text
                            let fullPath = Path.Combine(path, fileName)
                            let pdfDoc = PdfReader.Open(fullPath, PdfDocumentOpenMode.Import)
                            pdfSources.Value <- Array.append pdfSources.Value [| pdfDoc |]
                        label_Status.Text <- "Successfully Read Sources."
                        label_Status.Refresh()
                        label_Status.Text <- "Merging Files..."
                        label_Status.Refresh()
                        let pdfMerged = new PdfDocument()
                        for pd in pdfSources.Value do
                            for i=0 to pd.PageCount-1 do
                                pdfMerged.AddPage(pd.Pages.[i]) |> ignore
                        pdfMerged.Save(saveFileDialog.FileName)
                        label_Status.Text <- "Success!"
                        label_Status.Refresh()
                        MetroMessageBox.Show(mergeMultipleForm, "Successfully merged files.", "Success - Merge", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore
                        label_Status.Text <- "Ready"
                        label_Status.Refresh()
                    | DialogResult.Cancel ->
                        printfn "Cancelled merge."
                    | _ ->
                        printfn "%A" saveFileDialogResult
            
            (*if (textbox_InputPdf1.Text = "") then
                MetroMessageBox.Show(mergeMultipleForm, "Please select input source 1.", "Error - Unspecified source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (textbox_InputPdf2.Text = "") then
                MetroMessageBox.Show(mergeMultipleForm, "Please select input source 2.", "Error - Unspecified source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (PdfReader.TestPdfFile(textbox_InputPdf1.Text) = 0) then
                MetroMessageBox.Show(mergeMultipleForm, "Source 1 is invalid pdf file.", "Error - Invalid source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
            else if (PdfReader.TestPdfFile(textbox_InputPdf2.Text) = 0) then
                MetroMessageBox.Show(mergeMultipleForm, "Source 2 is invalid pdf file.", "Error - Invalid source", MessageBoxButtons.OK, MessageBoxIcon.Error) |> ignore
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
                    MetroMessageBox.Show(mergeMultipleForm, "Successfully merged files.", "Success - Merge", MessageBoxButtons.OK, MessageBoxIcon.Information) |> ignore
                    label_Status.Text <- "Ready"
                | DialogResult.Cancel ->
                    printfn "Cancelled merge."
                | _ ->
                    printfn "%A" saveFileDialogResult*)
        )

        mergeMultipleForm