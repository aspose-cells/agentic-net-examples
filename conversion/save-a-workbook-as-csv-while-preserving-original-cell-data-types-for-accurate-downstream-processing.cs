// Title: Save Aspose.Cells Workbook to CSV with DisplayStyle to Preserve Cell Formats
// Description: Shows how to export a worksheet to CSV using Aspose.Cells TxtSaveOptions with the DisplayStyle format strategy, keeping strings, numbers, dates, and booleans in their original displayed form.
// Keywords: Aspose.Cells CSV export | TxtSaveOptions | DisplayStyle format strategy | preserve cell formatting | C# Aspose.Cells example | save workbook as CSV | data type preservation | comma delimiter CSV | export single worksheet
// Common Searches: Aspose.Cells export to CSV preserving formats | DisplayStyle option for CSV in Aspose.Cells | keep date and number formatting when saving CSV with Aspose.Cells | C# save Excel as CSV without losing cell types | how to use TxtSaveOptions for CSV in Aspose.Cells
// Developer Intent: Export a worksheet to CSV while retaining the original display format of each cell.
// Use Cases: Create CSV reports from Excel files that contain mixed data types, ensuring downstream systems read the same textual values shown in Excel. | Migrate data from a spreadsheet to a CSV‑based analytics pipeline without losing numeric or date formatting. | Generate single‑sheet CSV files for integration with legacy applications that require exact string representations of dates, numbers, and booleans.
// AI Prompts: Write C# code that loads an existing Excel workbook and saves a selected worksheet to CSV using TxtSaveOptions with DisplayStyle to keep original formats. | Show how to change the CSV delimiter to a semicolon in TxtSaveOptions while still preserving cell display values. | Explain how to loop through all worksheets in a workbook and export each one to a separate CSV file, maintaining the displayed formatting for every sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsCsvExport
{
    // Shows how to export a worksheet to CSV using Aspose.Cells TxtSaveOptions with the DisplayStyle format strategy, keeping strings, numbers, dates, and booleans in their original displayed form.
    public class PreserveDataTypesCsv
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with different data types
                cells["A1"].PutValue("Text");                     // string
                cells["B1"].PutValue(12345);                      // integer
                cells["C1"].PutValue(123.456);                    // double
                cells["D1"].PutValue(DateTime.Now);               // DateTime
                cells["E1"].PutValue(true);                       // boolean

                // Configure CSV save options
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ',',                                 // Use comma as delimiter
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle, // Preserve displayed format
                    ExportAllSheets = false,                         // Export only the active sheet
                    ClearData = false                                // Keep workbook data after saving
                };

                // Save the workbook as CSV while preserving original data representations
                workbook.Save("PreservedDataTypes.csv", csvOptions);
                Console.WriteLine("CSV file saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PreserveDataTypesCsv.Run();
        }
    }
}
