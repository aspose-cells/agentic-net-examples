using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John, Doe");          // Contains delimiter
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("\"Alice\"");          // Contains text qualifier
            cells["B3"].PutValue(25);

            // Define the range to be exported (A1:B3)
            CellArea exportArea = new CellArea
            {
                StartRow = 0,   // Row index is zero‑based
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 1
            };

            // Configure TxtSaveOptions for CSV export
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                     // Delimiter to use
                QuoteType = TxtValueQuoteType.Always, // Always quote values to preserve qualifiers
                ExportArea = exportArea,              // Export only the defined range
                Encoding = Encoding.UTF8,             // Use UTF‑8 encoding
                TrimLeadingBlankRowAndColumn = true, // Optional: trim leading blanks
                KeepSeparatorsForBlankRow = false    // Optional: omit separators for empty rows
            };

            // Save the workbook as a CSV file using the configured options
            workbook.Save("ExportedRange.csv", saveOptions);

            Console.WriteLine("Range exported to CSV successfully.");
        }
    }
}