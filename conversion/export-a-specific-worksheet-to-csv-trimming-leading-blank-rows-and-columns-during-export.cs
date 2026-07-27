// Title: Export a specific worksheet to CSV with trimmed leading blank rows and columns – Aspose.Cells for .NET
// Description: Creates a workbook, adds a sheet named "DataSheet", fills data starting at C3, sets this sheet as active, and saves only the active sheet to a CSV file. The TxtSaveOptions are configured with TrimLeadingBlankRowAndColumn=true and ExportAllSheets=false, producing a CSV without any leading empty rows or columns.
// Keywords: Aspose.Cells CSV export | export active worksheet to CSV | TrimLeadingBlankRowAndColumn | TxtSaveOptions ExportAllSheets false | C# Aspose.Cells CSV trim blanks | save specific sheet as CSV | Aspose.Cells .NET conversion
// Common Searches: Aspose.Cells export active sheet to CSV | remove leading empty rows columns CSV Aspose | TxtSaveOptions TrimLeadingBlankRowAndColumn example C# | how to save one worksheet as CSV with Aspose.Cells | CSV export options Aspose.Cells .NET
// Developer Intent: Save only the selected worksheet as a CSV file while automatically removing any leading blank rows and columns.
// Use Cases: Generate a clean CSV report from a sheet that begins at C3, eliminating top‑left empty cells. | Automate the export of a single worksheet from a multi‑sheet workbook for downstream data pipelines. | Create CSV files compatible with systems that cannot process leading empty rows or columns.
// AI Prompts: Write C# code using Aspose.Cells to export the active worksheet to CSV, trimming leading blank rows and columns and exporting only that sheet. | Show how to configure TxtSaveOptions with TrimLeadingBlankRowAndColumn=true and ExportAllSheets=false for CSV conversion in Aspose.Cells. | Explain why setting a worksheet as active before saving to CSV matters and how trimming leading blanks improves data quality.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Creates a workbook, adds a sheet named "DataSheet", fills data starting at C3, sets this sheet as active, and saves only the active sheet to a CSV file. The TxtSaveOptions are configured with TrimLeadingBlankRowAndColumn=true and ExportAllSheets=false, producing a CSV without any leading empty rows or columns.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add a second worksheet to demonstrate selecting a specific sheet
            Worksheet sheet1 = workbook.Worksheets[0]; // default sheet
            Worksheet sheet2 = workbook.Worksheets.Add("DataSheet");

            // Populate sheet2 with data that has leading blank rows and columns
            // Row 0-1 and Column 0-1 are left blank intentionally
            sheet2.Cells["C3"].PutValue("First");
            sheet2.Cells["D4"].PutValue("Second");
            sheet2.Cells["E5"].PutValue("Third");

            // Make sheet2 the active worksheet (only the active sheet will be exported)
            workbook.Worksheets.ActiveSheetIndex = sheet2.Index;

            // Configure CSV (text) save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Trim leading blank rows and columns (default is true, set explicitly for clarity)
                TrimLeadingBlankRowAndColumn = true,
                // Export only the active worksheet (default false, keep as is)
                ExportAllSheets = false
            };

            // Save the active worksheet to CSV (lifecycle: save)
            workbook.Save("DataSheet_Trimmed.csv", saveOptions);
        }
    }
}
