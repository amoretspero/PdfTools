// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.
open MergePdf

[<EntryPoint>]
let main argv = 
    printfn "%A" argv
    
    let mainThread = new System.Threading.Thread(new System.Threading.ThreadStart(fun _ ->
        mergeForm.ShowDialog() |> ignore
    ))

    mainThread.SetApartmentState(System.Threading.ApartmentState.STA)

    mainThread.Start()

    0 // return an integer exit code
