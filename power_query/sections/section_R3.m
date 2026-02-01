// section_R3
// Orchestrates processing for section R.3 across all files in the input folder.
// Adds organization metadata via fnTitle and decodes the section via fnGetR3.

let
    // Read folder path from the current workbook named object 'link'
    FolderPath = Excel.CurrentWorkbook(){[Name="link"]}[Content]{0}[in],

    // Load all files from the folder
    Files = Folder.Files(FolderPath),

    // Build full file path
    WithFullPath = Table.AddColumn(Files, "FullPath", each [Folder Path] & [Name], type text),

    // Add organization name (metadata)
    WithOrgName = Table.AddColumn(WithFullPath, "OrganizationName", each fnTitle([FullPath])),

    // Decode R.3 (returns a single-row table per file)
    WithR3 = Table.AddColumn(WithOrgName, "R3", each fnGetR3([FullPath])),

    // Expand decoded metrics from the nested table
    Expanded = Table.ExpandTableColumn(
        WithR3,
        "R3",
        {
            "10001","10002",
            "11001","11002",
            "12001","12002",
            "12101","12102",
            "12201","12202",
            "12301","12302",
            "12401","12402",
            "12411","12412",
            "20001","20002",
            "21001","21002"
        },
        {
            "10001","10002",
            "11001","11002",
            "12001","12002",
            "12101","12102",
            "12201","12202",
            "12301","12302",
            "12401","12402",
            "12411","12412",
            "20001","20002",
            "21001","21002"
        }
    ),

    // Keep only essential fields + decoded metrics
    Cleaned = Table.RemoveColumns(
        Expanded,
        {"Content","Extension","Date accessed","Date modified","Date created","Attributes","Folder Path","FullPath"}
    )
in
    Cleaned
