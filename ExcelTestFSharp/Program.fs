// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.

open FSharp.ExcelProvider
open FSharp.Data
open Microsoft.FSharp.Collections

type ExcelData = ExcelFile<"Data\part1.xlsx", ForceString=true>
type DexXml = XmlProvider<Schema="Data\DEXFileUpload (1).xsd">

type Row = ExcelData.Row

let parseBool s =
    match s with
    | "true" | "True" | "TRUE" -> true
    | _ -> false

let parseDate s =
    System.DateTime.Parse s

let parseSlk s =
    Some (new DexXml.Slk(
                None,
                Some s))

let getAddress (row:Row) =
    new DexXml.ResidentialAddress(
            None,
            None,
            row.Suburb,
            row.StateCode,
            row.Postcode)

let getClient (row:Row) =
    new DexXml.Client(
            row.ClientId,
            parseSlk(row.Slk),
            parseBool(row.ConsentToProvideDetails),
            parseBool(row.ConsentedForFutureContacts),
            None,
            None,
            parseBool(row.IsUsingPsuedonym),
            parseDate(row.BirthDate),
            parseBool(row.IsBirthDateAnEstimate),
            row.GenderCode,
            row.CountryOfBirthCode,
            row.LanguageSpokenAtHomeCode,
            row.AboriginalOrTorresStraitIslanderOriginCode,
            parseBool(row.HasDisabilities),
            None,
            None,
            None,
            None,
            getAddress row,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None)

let getReason (row:Row) =
    new DexXml.ReasonForAssistance(
            row.ReasonForAssistanceCode,
            parseBool(row.IsPrimary))

let getCaseClient (row:Row) =
    new DexXml.CaseClient(
        row.ClientId.ToString(),
        None,
        Some(new DexXml.ReasonsForAssistance(
                    None,
                    [| getReason row |])),
        None)

let getCase (row:Row) =
    new DexXml.Case(
            row.CaseId.ToString(),
            row.OutletActivityId.ToString(),
            row.TotalNumberOfUnidentifiedClients.ToString(),
            Some(new DexXml.CaseClients(
                        None,
                        [| getCaseClient row |])),
            None,
            None)

let getParts (row:Row) =
    getClient row, getCase row
        
[<EntryPoint>]
let main argv = 
    let file = new ExcelData()
    
    let (clientArray, caseArray) = 
        file.Data
        |> Seq.filter(fun x -> x.ClientId <> null)
        |> Seq.map getParts
        |> Seq.toArray
        |> Array.unzip
    
    let clients = [| new DexXml.Clients(clientArray) |]
    let cases = [| new DexXml.Cases(caseArray) |]

    let output = new DexXml.DexFileUpload(clients, cases, [||], [||], [||])
    
    output.XElement.Save("output.xml")
    0
