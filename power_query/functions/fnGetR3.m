// fnGetR3(filepath)
// Extracts and decodes section 'Р.3' from a report Excel file.
// Pattern: remove noise → skip headers → transpose → promote headers → expand coded metrics
// Output: single-row, wide-format table

(filepath as text) as table =>
let
    // --- Load workbook and sheet ---
    Workbook = Excel.Workbook(File.Contents(filepath), null, true),
    Sheet = Workbook{[Item="Р.3", Kind="Sheet"]}[Data],

    // --- Basic cleanup ---
    RemovedColumns = Table.RemoveColumns(Sheet, {"Column1", "Column5", "Column6", "Column7"}),

    // --- Remove header rows ---
    SkippedHeaderRows = Table.Skip(RemovedColumns, 3),

    // --- Matrix to table ---
    Transposed = Table.Transpose(SkippedHeaderRows),
    PromotedHeaders = Table.PromoteHeaders(Transposed, [PromoteAllScalars = true]),

    // --- Expand coded metrics (2 positions per base code in this section) ---
    Add10001 = Table.AddColumn(PromotedHeaders, "10001", each PromotedHeaders{0}[1000]),
    Add10002 = Table.AddColumn(Add10001,       "10002", each PromotedHeaders{1}[1000]),

    Add11001 = Table.AddColumn(Add10002,       "11001", each PromotedHeaders{0}[1100]),
    Add11002 = Table.AddColumn(Add11001,       "11002", each PromotedHeaders{1}[1100]),

    Add12001 = Table.AddColumn(Add11002,       "12001", each PromotedHeaders{0}[1200]),
    Add12002 = Table.AddColumn(Add12001,       "12002", each PromotedHeaders{1}[1200]),

    Add12101 = Table.AddColumn(Add12002,       "12101", each PromotedHeaders{0}[1210]),
    Add12102 = Table.AddColumn(Add12101,       "12102", each PromotedHeaders{1}[1210]),

    Add12201 = Table.AddColumn(Add12102,       "12201", each PromotedHeaders{0}[1220]),
    Add12202 = Table.AddColumn(Add12201,       "12202", each PromotedHeaders{1}[1220]),

    Add12301 = Table.AddColumn(Add12202,       "12301", each PromotedHeaders{0}[1230]),
    Add12302 = Table.AddColumn(Add12301,       "12302", each PromotedHeaders{1}[1230]),

    Add12401 = Table.AddColumn(Add12302,       "12401", each PromotedHeaders{0}[1240]),
    Add12402 = Table.AddColumn(Add12401,       "12402", each PromotedHeaders{1}[1240]),

    Add12411 = Table.AddColumn(Add12402,       "12411", each PromotedHeaders{0}[1241]),
    Add12412 = Table.AddColumn(Add12411,       "12412", each PromotedHeaders{1}[1241]),

    Add20001 = Table.AddColumn(Add12412,       "20001", each PromotedHeaders{0}[2000]),
    Add20002 = Table.AddColumn(Add20001,       "20002", each PromotedHeaders{1}[2000]),

    Add21001 = Table.AddColumn(Add20002,       "21001", each PromotedHeaders{0}[2100]),
    Add21002 = Table.AddColumn(Add21001,       "21002", each PromotedHeaders{1}[2100]),

    // --- Remove base-code columns ---
    RemovedBaseCodes = Table.RemoveColumns(
        Add21002,
        {"Б","1000","1100","1200","1210","1220","1230","1240","1241","2000","2100"}
    ),

    // --- Ensure single row output ---
    DistinctRow = Table.Distinct(RemovedBaseCodes)
in
    DistinctRow
in
    fnGetR3
