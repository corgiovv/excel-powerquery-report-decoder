// section_R11
// Orchestrates processing for section R.1.1 across all files in the input folder.
// Steps:
// 1) Load files from folder (path stored in named object 'link')
// 2) Extract organization name via fnTitle
// 3) Decode section R.1.1 via fnGetR11
// 4) Expand decoded columns and remove technical metadata

let
    // Read folder path from the current workbook named object 'link'
    FolderPath = Excel.CurrentWorkbook(){[Name="link"]}[Content]{0}[in],

    // Load all files from the folder
    Files = Folder.Files(FolderPath),

    // Build full file path
    WithFullPath = Table.AddColumn(Files, "FullPath", each [Folder Path] & [Name], type text),

    // Add organization name (metadata)
    WithOrgName = Table.AddColumn(WithFullPath, "OrganizationName", each fnTitle([FullPath])),

    // Decode R.1.1 (returns a single-row table per file)
    WithR11 = Table.AddColumn(WithOrgName, "R11", each fnGetR11([FullPath])),

    // Expand decoded metrics from the nested table
    Expanded = Table.ExpandTableColumn(
        WithR11,
        "R11",
        {
            "10001","10002","10003","10004",
            "11001","11002","11003","11004",
            "12001","12002","12003","12004",
            "21001","21002","21003","21004",
            "22001","22002","22003","22004",
            "22101","22102","22103","22104",
            "31001","31002","31003","31004",
            "32001","32002","32003","32004",
            "41001","41002","41003","41004",
            "42001","42002","42003","42004",
            "51001","51002","51003","51004",
            "52001","52002","52003","52004",
            "52101","52102","52103","52104",
            "61001","61002","61003","61004",
            "62001","62002","62003","62004"
        },
        {
            "10001","10002","10003","10004",
            "11001","11002","11003","11004",
            "12001","12002","12003","12004",
            "21001","21002","21003","21004",
            "22001","22002","22003","22004",
            "22101","22102","22103","22104",
            "31001","31002","31003","31004",
            "32001","32002","32003","32004",
            "41001","41002","41003","41004",
            "42001","42002","42003","42004",
            "51001","51002","51003","51004",
            "52001","52002","52003","52004",
            "52101","52102","52103","52104",
            "61001","61002","61003","61004",
            "62001","62002","62003","62004"
        }
    ),

    // Keep only the essential fields + decoded metrics
    Cleaned = Table.RemoveColumns(
        Expanded,
        {"Content","Extension","Date accessed","Date modified","Date created","Attributes","Folder Path","FullPath"}
    )
in
    Cleaned
