// Title: Export Worksheet to CSV as Text with Aspose.Cells (C#)
// Description: The example creates a workbook, populates cells with various data types, and saves the active worksheet as a CSV file using TxtSaveOptions configured to output the displayed value, quote every field, and use UTF‑8 encoding, ensuring each cell is written as plain text.
// Keywords: Aspose.Cells | CSV export | C# | TxtSaveOptions | DisplayStyle | QuoteType.Always | UTF-8 | preserve formatting | export active sheet | cell values as text
// Common Searches: Aspose.Cells export CSV as text | C# save worksheet to CSV with quoted fields | keep numeric and date formatting in CSV Aspose.Cells | export only active sheet to CSV Aspose.Cells | set UTF-8 encoding for CSV export Aspose.Cells
// Developer Intent: Create a CSV file from a worksheet where every cell is emitted as a text string, retaining the cell’s displayed formatting.
// Use Cases: Generate CSV files for downstream systems that require all columns quoted and interpreted as strings. | Produce UTF‑8 encoded CSV reports while preserving numeric, date, and boolean display formats. | Export only the active worksheet to CSV, omitting other sheets in the workbook. | Convert Excel data to a text‑only CSV for import into databases or legacy applications.
// AI Prompts: Show how to export multiple worksheets to separate CSV files, each with all values quoted as text. | Explain the difference between CellValueFormatStrategy.RawValue and DisplayStyle when saving CSV with Aspose.Cells. | Provide C# code to use a custom delimiter (e.g., semicolon) while still quoting every field in the CSV output.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // The example creates a workbook, populates cells with various data types, and saves the active worksheet as a CSV file using TxtSaveOptions configured to output the displayed value, quote every field, and use UTF‑8 encoding, ensuring each cell is written as plain text.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with different data types
            cells["A1"].PutValue("Text");
            cells["B1"].PutValue(123);               // Integer
            cells["A2"].PutValue(45.67);             // Double
            cells["B2"].PutValue(DateTime.Now);      // DateTime
            cells["A3"].PutValue(true);              // Boolean

            // Create TxtSaveOptions for CSV format
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Export the displayed value as text (preserves formatting)
                FormatStrategy = CellValueFormatStrategy.DisplayStyle,
                // Quote all fields to ensure they are treated as text when read back
                QuoteType = TxtValueQuoteType.Always,
                // Use UTF-8 encoding
                Encoding = Encoding.UTF8,
                // Export only the active sheet (default)
                ExportAllSheets = false
            };

            // Save the worksheet as CSV with the specified options
            workbook.Save("ExportedWorksheet.csv", saveOptions);
        }
    }
}
