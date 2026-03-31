using System;
using System.Data;
using Aspose.Cells;

class PreserveNonStronglyTypedData
{
    static void Main()
    {
        // Load the source workbook that contains mixed (non‑strongly typed) data
        Workbook sourceWb = new Workbook("source.xlsx");
        Worksheet sourceWs = sourceWb.Worksheets[0];

        // Determine the used range size
        int totalRows = sourceWs.Cells.MaxDataRow + 1;      // zero‑based index + 1 = count
        int totalCols = sourceWs.Cells.MaxDataColumn + 1;   // zero‑based index + 1 = count

        // Export the whole used range as a DataTable where every cell is treated as a string.
        // This preserves the original textual representation (e.g., "00123", "12/31/2023", "ABC").
        DataTable rawTable = sourceWs.Cells.ExportDataTableAsString(0, 0, totalRows, totalCols, false);

        // Create a new workbook where we will import the data without any automatic conversion.
        Workbook targetWb = new Workbook();
        Worksheet targetWs = targetWb.Worksheets[0];

        // Configure import options: do NOT convert strings to numbers or dates.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            ConvertNumericData = false,   // keep everything as the original string
            IsFieldNameShown = false,     // do not treat the first row as headers
            InsertRows = true             // add rows if needed
        };

        // Import the raw DataTable into the target worksheet, preserving original values.
        targetWs.Cells.ImportData(rawTable, 0, 0, importOptions);

        // Example manipulation: change the value of the first cell while keeping other cells untouched.
        // The new value can be any type; here we replace it with a string.
        targetWs.Cells[0, 0].PutValue("ModifiedValue");

        // Save the resulting workbook.
        targetWb.Save("preserved_output.xlsx");
    }
}