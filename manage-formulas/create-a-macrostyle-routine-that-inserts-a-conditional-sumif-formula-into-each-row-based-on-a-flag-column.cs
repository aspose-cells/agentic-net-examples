// Title: Programmatically add a conditional SUMIF formula to each row based on a flag column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add SUMIF formula to each row using a flag column | Insert conditional SUMIF in Excel worksheet with Aspose.Cells .NET | C# loop to set SUMIF formula across rows in Aspose.Cells workbook | Generate dynamic range for SUMIF in Aspose.Cells based on last data row | Macro‑style routine for flag‑based aggregation in Excel using Aspose.Cells | How to apply SUMIF only when flag equals Yes in Aspose.Cells
// Tags: SUMIF formula insertion Aspose.Cells | C# loop set Excel formula per row | dynamic range generation Aspose.Cells | flag column based SUMIF .NET | macro style Excel automation Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, populates sample data, determines the last used row, builds absolute ranges for the amount and flag columns, and loops through each data row to assign a SUMIF formula in column D that sums amounts only when the flag column equals "Yes". The workbook is then saved as ConditionalSum.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set up headers (optional, for clarity)
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["C1"].PutValue("Flag");
        sheet.Cells["D1"].PutValue("ConditionalSum");

        // -------------------------------------------------
        // Example data population (can be omitted in real use)
        // -------------------------------------------------
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["C2"].PutValue("Yes");

        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["C3"].PutValue("No");

        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["C4"].PutValue("Yes");
        // -------------------------------------------------

        // Determine the last used row (zero‑based index)
        int lastRowIndex = sheet.Cells.MaxDataRow; // includes header row

        // Build absolute range strings for SUMIF
        // Excel rows are 1‑based, so add 1 to the zero‑based index
        string amountRange = $"$B$2:$B${lastRowIndex + 1}";
        string flagRange   = $"$C$2:$C${lastRowIndex + 1}";
        string criteria    = "\"Yes\""; // flag value that triggers inclusion

        // Insert the conditional SUMIF formula into each data row (column D)
        // Row index starts at 1 (Excel row 2) to skip the header
        for (int rowIdx = 1; rowIdx <= lastRowIndex; rowIdx++)
        {
            // Formula: =SUMIF($C$2:$C${lastRow}, "Yes", $B$2:$B${lastRow})
            string formula = $"SUMIF({flagRange}, {criteria}, {amountRange})";
            sheet.Cells[rowIdx, 3].Formula = formula; // Column D has index 3 (0‑based)
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("ConditionalSum.xlsx");
    }
}
