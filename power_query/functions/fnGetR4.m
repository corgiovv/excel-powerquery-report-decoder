// fnGetR4(filepath)
// Extracts and decodes section 'Р.4' from a report Excel file.
// Pattern: remove noise → skip headers → remove footer → transpose → promote headers → expand coded metrics
// Output: single-row, wide-format table

(filepath as text) as table =>
let
    // --- Load workbook and sheet ---
    Workbook = Excel.Workbook(File.Contents(filepath), null, true),
    Sheet = Workbook{[Item="Р.4", Kind="Sheet"]}[Data],

    // --- Basic cleanup ---
    RemovedColumns = Table.RemoveColumns(Sheet, {"Column1"}),

    // --- Remove header + footer template rows ---
    SkippedHeaderRows = Table.Skip(RemovedColumns, 3),
    RemovedFooterRows = Table.RemoveLastN(SkippedHeaderRows, 10),

    // --- Matrix to table ---
    Transposed = Table.Transpose(RemovedFooterRows),
    PromotedHeaders = Table.PromoteHeaders(Transposed, [PromoteAllScalars = true]),

    // --- Expand coded metrics (2 positions per base code in this section) ---
    Add10001 = Table.AddColumn(PromotedHeaders, "10001", each PromotedHeaders{0}[1000]),
    Add10002 = Table.AddColumn(Add10001,       "10002", each PromotedHeaders{1}[1000]),

    Add20001 = Table.AddColumn(Add10002,       "20001", each PromotedHeaders{0}[2000]),
    Add20002 = Table.AddColumn(Add20001,       "20002", each PromotedHeaders{1}[2000]),

    // --- Remove base-code columns ---
    RemovedBaseCodes = Table.RemoveColumns(Add20002, {"Б","1000","2000"}),

    // --- Ensure single row output ---
    DistinctRow = Table.Distinct(RemovedBaseCodes)
in
    DistinctRow
in
    fnGetR4
