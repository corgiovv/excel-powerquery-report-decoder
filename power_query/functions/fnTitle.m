// fnTitle(filepath)
// Extracts organization name from the 'title' sheet of a report file
// The value is read from a fixed cell position in the template

(filepath as text) as any =>
let
    // Load Excel workbook
    Workbook = Excel.Workbook(File.Contents(filepath), null, true),

    // Access 'title' sheet
    TitleSheet = Workbook{[Item="title", Kind="Sheet"]}[Data],

    // Read organization name from a fixed cell
    OrganizationName = TitleSheet{21}[Column2]
in
    OrganizationName
