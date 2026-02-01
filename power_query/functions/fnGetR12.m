// fnGetR12(filepath)
// Extracts and decodes section 'Р.1.2' from a report Excel file.
// Pattern: cleanup → skip headers → transpose → promote code headers → expand coded metrics
// Output: single-row, wide-format table

(filepath as text) as table =>
let
    // --- Load workbook and sheet ---
    Workbook = Excel.Workbook(File.Contents(filepath), null, true),
    Sheet = Workbook{[Item="Р.1.2", Kind="Sheet"]}[Data],

    // --- Basic cleanup ---
    ChangedTypes = Table.TransformColumnTypes(
        Sheet,
        {
            {"Column1", type text},
            {"Column2", type any},
            {"Column3", type any},
            {"Column4", type any},
            {"Column5", type any},
            {"Column6", type any},
            {"Column7", type any}
        }
    ),
    RemovedTextColumn = Table.RemoveColumns(ChangedTypes, {"Column1"}),

    // --- Remove header rows ---
    SkippedHeaderRows = Table.Skip(RemovedTextColumn, 4),

    // --- Matrix to table ---
    Transposed = Table.Transpose(SkippedHeaderRows),
    PromotedHeaders = Table.PromoteHeaders(Transposed, [PromoteAllScalars = true]),

    // --- Expand coded metrics (5 positions per base code in this section) ---
    Add10001 = Table.AddColumn(PromotedHeaders, "10001", each PromotedHeaders{0}[1000]),
    Add10002 = Table.AddColumn(Add10001,       "10002", each PromotedHeaders{1}[1000]),
    Add10003 = Table.AddColumn(Add10002,       "10003", each PromotedHeaders{2}[1000]),
    Add10004 = Table.AddColumn(Add10003,       "10004", each PromotedHeaders{3}[1000]),
    Add10005 = Table.AddColumn(Add10004,       "10005", each PromotedHeaders{4}[1000]),

    Add11001 = Table.AddColumn(Add10005,       "11001", each PromotedHeaders{0}[1100]),
    Add11002 = Table.AddColumn(Add11001,       "11002", each PromotedHeaders{1}[1100]),
    Add11003 = Table.AddColumn(Add11002,       "11003", each PromotedHeaders{2}[1100]),
    Add11004 = Table.AddColumn(Add11003,       "11004", each PromotedHeaders{3}[1100]),
    Add11005 = Table.AddColumn(Add11004,       "11005", each PromotedHeaders{4}[1100]),

    Add11101 = Table.AddColumn(Add11005,       "11101", each PromotedHeaders{0}[1110]),
    Add11102 = Table.AddColumn(Add11101,       "11102", each PromotedHeaders{1}[1110]),
    Add11103 = Table.AddColumn(Add11102,       "11103", each PromotedHeaders{2}[1110]),
    Add11104 = Table.AddColumn(Add11103,       "11104", each PromotedHeaders{3}[1110]),
    Add11105 = Table.AddColumn(Add11104,       "11105", each PromotedHeaders{4}[1110]),

    // --- Remove base-code columns ---
    RemovedBaseCodes = Table.RemoveColumns(Add11105, {"Б","1000","1100","1110"}),

    // --- Ensure single row output ---
    DistinctRow = Table.Distinct(RemovedBaseCodes, {"10001"})
in
    DistinctRow
in
    fnGetR12
