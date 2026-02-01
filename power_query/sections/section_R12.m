// section_R12
// Orchestrates processing for section R.1.2 across all files in the input folder.
// Adds organization metadata via fnTitle and decodes the section via fnGetR12.

let
    // Read folder path from the current workbook named object 'link'
    FolderPath = Excel.CurrentWorkbook(){[Name="link"]}[Content]{0}[in],

    // Load all files from the folder
    Files = Folder.Files(FolderPath),

    // Build full file path
    WithFullPath = Table.AddColumn(Files, "FullPath", each [Folder Path] & [Name], type text),

    // Add organization name (metadata)
    WithOrgName = Table.AddColumn(WithFullPath, "OrganizationName", each fnTitle([FullPath])),

    // Decode R.1.2 (returns a single-row table per file)
    WithR12 = Table.AddColumn(WithOrgName, "R12", each fnGetR12([FullPath])),

    // Expand decoded metrics from the nested table
    Expanded = Table.ExpandTableColumn(
        WithR12,
        "R12",
        {
            "10001","10002","10003","10004","10005",
            "11001","11002","11003","11004","11005",
            "11101","11102","11103","11104","11105"
        },
        {
            "10001","10002","10003","10004","10005",
            "11001","11002","11003","11004","11005",
            "11101","11102","11103","11104","11105"
        }
    ),

    // Type percentage fields (as in the original workflow)
    Typed = Table.TransformColumnTypes(
        Expanded,
        {{"10005", Percentage.Type}, {"11005", Percentage.Type}, {"11105", Percentage.Type}}
    ),

    // Keep only essential fields + decoded metrics
    Cleaned = Table.RemoveColumns(
        Typed,
        {"Content","Extension","Date accessed","Date modified","Date created","Attributes","Folder Path","FullPath"}
    )
in
    Cleaned
