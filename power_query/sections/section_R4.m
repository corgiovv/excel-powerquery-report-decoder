// section_R4
// Orchestrates processing for section R.4 across all files in the input folder.
// Adds organization metadata via fnTitle and decodes the section via fnGetR4.

let
    // Read folder path from the current workbook named object 'link'
    FolderPath = Excel.CurrentWorkbook(){[Name="link"]}[Content]{0}[in],

    // Load all files from the folder
    Files = Folder.Files(FolderPath),

    // Build full file path
    WithFullPath = Table.AddColumn(Files, "FullPath", each [Folder Path] & [Name], type text),

    // Add organization name (metadata)
    WithOrgName = Table.AddColumn(WithFullPath, "OrganizationName", each fnTitle([FullPath])),

    // Decode R.4 (returns a single-row table per file)
    WithR4 = Table.AddColumn(WithOrgName, "R4", each fnGetR4([FullPath])),

    // Expand decoded metrics from the nested table
    Expanded = Table.ExpandTableColumn(
        WithR4,
        "R4",
        {"10001","10002","20001","20002"},
        {"10001","10002","20001","20002"}
    ),

    // Keep only essential fields + decoded metrics
    Cleaned = Table.RemoveColumns(
        Expanded,
        {"Content","Extension","Date accessed","Date modified","Date created","Attributes","Folder Path","FullPath"}
    )
in
    Cleaned
