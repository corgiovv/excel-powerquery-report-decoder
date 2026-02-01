// fnGetR11(filepath)
// Extracts and decodes section 'Р.1.1' from a report Excel file.
// The template is semi-structured, so the function:
// 1) loads the sheet
// 2) removes non-data columns and header rows
// 3) transposes the matrix
// 4) promotes code headers (e.g., 1000, 1100, ...)
// 5) expands codes into dedicated output columns (e.g., 10001..10004)
// Output: single-row, wide-format table

(filepath as text) as table =>
let
    // --- Load workbook and sheet ---
    Workbook = Excel.Workbook(File.Contents(filepath), null, true),
    Sheet = Workbook{[Item="Р.1.1", Kind="Sheet"]}[Data],

    // --- Basic cleanup ---
    ChangedTypes = Table.TransformColumnTypes(
        Sheet,
        {
            {"Column1", type any},
            {"Column2", type text},
            {"Column3", type any},
            {"Column4", type any},
            {"Column5", type any},
            {"Column6", type any},
            {"Column7", type any}
        }
    ),
    RemovedTextColumns = Table.RemoveColumns(ChangedTypes, {"Column1", "Column2"}),

    // --- Remove header rows ---
    SkippedHeaderRows = Table.Skip(RemovedTextColumns, 6),

    // --- Matrix to table ---
    Transposed = Table.Transpose(SkippedHeaderRows),
    PromotedHeaders = Table.PromoteHeaders(Transposed, [PromoteAllScalars = true]),

    // --- Expand coded metrics (4 positions per base code) ---
    Add10001 = Table.AddColumn(PromotedHeaders, "10001", each PromotedHeaders{0}[1000]),
    Add10002 = Table.AddColumn(Add10001,       "10002", each PromotedHeaders{1}[1000]),
    Add10003 = Table.AddColumn(Add10002,       "10003", each PromotedHeaders{2}[1000]),
    Add10004 = Table.AddColumn(Add10003,       "10004", each PromotedHeaders{3}[1000]),

    Add11001 = Table.AddColumn(Add10004,       "11001", each PromotedHeaders{0}[1100]),
    Add11002 = Table.AddColumn(Add11001,       "11002", each PromotedHeaders{1}[1100]),
    Add11003 = Table.AddColumn(Add11002,       "11003", each PromotedHeaders{2}[1100]),
    Add11004 = Table.AddColumn(Add11003,       "11004", each PromotedHeaders{3}[1100]),

    Add12001 = Table.AddColumn(Add11004,       "12001", each PromotedHeaders{0}[1200]),
    Add12002 = Table.AddColumn(Add12001,       "12002", each PromotedHeaders{1}[1200]),
    Add12003 = Table.AddColumn(Add12002,       "12003", each PromotedHeaders{2}[1200]),
    Add12004 = Table.AddColumn(Add12003,       "12004", each PromotedHeaders{3}[1200]),

    Add21001 = Table.AddColumn(Add12004,       "21001", each PromotedHeaders{0}[2100]),
    Add21002 = Table.AddColumn(Add21001,       "21002", each PromotedHeaders{1}[2100]),
    Add21003 = Table.AddColumn(Add21002,       "21003", each PromotedHeaders{2}[2100]),
    Add21004 = Table.AddColumn(Add21003,       "21004", each PromotedHeaders{3}[2100]),

    Add22001 = Table.AddColumn(Add21004,       "22001", each PromotedHeaders{0}[2200]),
    Add22002 = Table.AddColumn(Add22001,       "22002", each PromotedHeaders{1}[2200]),
    Add22003 = Table.AddColumn(Add22002,       "22003", each PromotedHeaders{2}[2200]),
    Add22004 = Table.AddColumn(Add22003,       "22004", each PromotedHeaders{3}[2200]),

    Add22101 = Table.AddColumn(Add22004,       "22101", each PromotedHeaders{0}[2210]),
    Add22102 = Table.AddColumn(Add22101,       "22102", each PromotedHeaders{1}[2210]),
    Add22103 = Table.AddColumn(Add22102,       "22103", each PromotedHeaders{2}[2210]),
    Add22104 = Table.AddColumn(Add22103,       "22104", each PromotedHeaders{3}[2210]),

    Add31001 = Table.AddColumn(Add22104,       "31001", each PromotedHeaders{0}[3100]),
    Add31002 = Table.AddColumn(Add31001,       "31002", each PromotedHeaders{1}[3100]),
    Add31003 = Table.AddColumn(Add31002,       "31003", each PromotedHeaders{2}[3100]),
    Add31004 = Table.AddColumn(Add31003,       "31004", each PromotedHeaders{3}[3100]),

    Add32001 = Table.AddColumn(Add31004,       "32001", each PromotedHeaders{0}[3200]),
    Add32002 = Table.AddColumn(Add32001,       "32002", each PromotedHeaders{1}[3200]),
    Add32003 = Table.AddColumn(Add32002,       "32003", each PromotedHeaders{2}[3200]),
    Add32004 = Table.AddColumn(Add32003,       "32004", each PromotedHeaders{3}[3200]),

    Add41001 = Table.AddColumn(Add32004,       "41001", each PromotedHeaders{0}[4100]),
    Add41002 = Table.AddColumn(Add41001,       "41002", each PromotedHeaders{1}[4100]),
    Add41003 = Table.AddColumn(Add41002,       "41003", each PromotedHeaders{2}[4100]),
    Add41004 = Table.AddColumn(Add41003,       "41004", each PromotedHeaders{3}[4100]),

    Add42001 = Table.AddColumn(Add41004,       "42001", each PromotedHeaders{0}[4200]),
    Add42002 = Table.AddColumn(Add42001,       "42002", each PromotedHeaders{1}[4200]),
    Add42003 = Table.AddColumn(Add42002,       "42003", each PromotedHeaders{2}[4200]),
    Add42004 = Table.AddColumn(Add42003,       "42004", each PromotedHeaders{3}[4200]),

    Add51001 = Table.AddColumn(Add42004,       "51001", each PromotedHeaders{0}[5100]),
    Add51002 = Table.AddColumn(Add51001,       "51002", each PromotedHeaders{1}[5100]),
    Add51003 = Table.AddColumn(Add51002,       "51003", each PromotedHeaders{2}[5100]),
    Add51004 = Table.AddColumn(Add51003,       "51004", each PromotedHeaders{3}[5100]),

    Add52001 = Table.AddColumn(Add51004,       "52001", each PromotedHeaders{0}[5200]),
    Add52002 = Table.AddColumn(Add52001,       "52002", each PromotedHeaders{1}[5200]),
    Add52003 = Table.AddColumn(Add52002,       "52003", each PromotedHeaders{2}[5200]),
    Add52004 = Table.AddColumn(Add52003,       "52004", each PromotedHeaders{3}[5200]),

    Add52101 = Table.AddColumn(Add52004,       "52101", each PromotedHeaders{0}[5210]),
    Add52102 = Table.AddColumn(Add52101,       "52102", each PromotedHeaders{1}[5210]),
    Add52103 = Table.AddColumn(Add52102,       "52103", each PromotedHeaders{2}[5210]),
    Add52104 = Table.AddColumn(Add52103,       "52104", each PromotedHeaders{3}[5210]),

    Add61001 = Table.AddColumn(Add52104,       "61001", each PromotedHeaders{0}[6100]),
    Add61002 = Table.AddColumn(Add61001,       "61002", each PromotedHeaders{1}[6100]),
    Add61003 = Table.AddColumn(Add61002,       "61003", each PromotedHeaders{2}[6100]),
    Add61004 = Table.AddColumn(Add61003,       "61004", each PromotedHeaders{3}[6100]),

    Add62001 = Table.AddColumn(Add61004,       "62001", each PromotedHeaders{0}[6200]),
    Add62002 = Table.AddColumn(Add62001,       "62002", each PromotedHeaders{1}[6200]),
    Add62003 = Table.AddColumn(Add62002,       "62003", each PromotedHeaders{2}[6200]),
    Add62004 = Table.AddColumn(Add62003,       "62004", each PromotedHeaders{3}[6200]),

    // --- Remove base-code columns, keep only decoded columns ---
    RemovedBaseCodes = Table.RemoveColumns(
        Add62004,
        {"1000","1100","1200","2000","2100","2200","2210","3000","3100","3200","4000","4100","4200","5000","5100","5200","5210","6000","6100","6200"}
    ),

    // --- Ensure single row output ---
    DistinctRow = Table.Distinct(RemovedBaseCodes, {"10001"})
in
    DistinctRow
in
    fnGetR11
