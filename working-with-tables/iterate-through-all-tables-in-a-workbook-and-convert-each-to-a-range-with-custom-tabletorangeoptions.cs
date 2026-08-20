// Title: Batch Convert Excel Tables to Ranges with TableToRangeOptions (Aspose.Cells C#)
// Description: Loads a workbook, loops through every worksheet and each ListObject (Excel table), creates a TableToRangeOptions object with the table's EndRow, converts the table to a plain range, and saves the updated file. Ideal for bulk table‑to‑range transformations while keeping data intact.
// Keywords: Aspose.Cells convert table to range C# | TableToRangeOptions example | ListObject to range Aspose.Cells | batch table conversion Excel C# | iterate worksheets tables Aspose | Excel table to plain range | Aspose.Cells LastRow option | C# Aspose.Cells table conversion
// Common Searches: Aspose.Cells convert all tables to ranges | TableToRangeOptions LastRow C# example | Iterate worksheets and ListObjects Aspose.Cells | Batch convert Excel tables using Aspose | How to change Excel tables to ranges in .NET
// Developer Intent: Programmatically change every ListObject in a workbook into a regular range using custom TableToRangeOptions.
// Use Cases: Strip table formatting before exporting to CSV or PDF. | Prepare workbooks for third‑party engines that only recognize plain ranges. | Apply a uniform conversion rule (e.g., fixing the last row) after dynamically resizing tables.
// AI Prompts: Write C# code that logs each table name while converting it to a range with TableToRangeOptions. | Show how to preserve cell styles and formulas when using TableToRangeOptions in a batch conversion. | Explain additional TableToRangeOptions properties such as FirstRow, PreserveFormatting, and how to combine them in a loop.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads a workbook, loops through every worksheet and each ListObject (Excel table), creates a TableToRangeOptions object with the table's EndRow, converts the table to a plain range, and saves the updated file. Ideal for bulk table‑to‑range transformations while keeping data intact.
class ConvertTablesToRanges
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through each table (ListObject) in the current worksheet
            foreach (ListObject table in worksheet.ListObjects)
            {
                // Create conversion options and set the last row to the table's current end row
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = table.EndRow
                };

                // Convert the table to a regular range using the specified options
                table.ConvertToRange(options);
            }
        }

        // Save the workbook after all tables have been converted
        workbook.Save("output.xlsx");
    }
}
