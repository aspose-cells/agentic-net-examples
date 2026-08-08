// Title: Export Active Worksheet to Semicolon‑Delimited CSV with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, marks the first sheet as active, configures TxtSaveOptions with SaveFormat.Csv and a ';' separator, disables ExportAllSheets, and saves only the active sheet to a CSV file.
// Keywords: Aspose.Cells | C# | .NET | CSV export | semicolon delimiter | TxtSaveOptions | SaveFormat.Csv | active worksheet | single sheet CSV | European CSV format
// Common Searches: Aspose.Cells export active sheet to CSV | CSV separator semicolon Aspose.Cells .NET | Save only one worksheet as CSV using Aspose.Cells | How to set custom delimiter when saving CSV with Aspose.Cells | Export workbook sheet to CSV with ';' delimiter
// Developer Intent: Save only the active worksheet of a workbook as a CSV file using a semicolon as the field delimiter.
// Use Cases: Generate a semicolon‑delimited CSV report from the currently active sheet in a newly created workbook. | Produce CSV files for European locales where the semicolon is the default list separator. | Export a single sheet from a multi‑sheet workbook without writing the other sheets to the CSV output.
// AI Prompts: Show how to modify the code to export the active sheet to a custom file path with UTF‑8 encoding using Aspose.Cells. | Provide an example that reads a semicolon‑delimited CSV into a workbook, sets a specific worksheet as active, and then saves that sheet back to CSV. | Explain how to export multiple worksheets each to separate semicolon‑delimited CSV files in a loop with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Creates a workbook, adds sample data, marks the first sheet as active, configures TxtSaveOptions with SaveFormat.Csv and a ';' separator, disables ExportAllSheets, and saves only the active sheet to a CSV file.
    public class ExportActiveSheetToCsv
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // Access the first (active) worksheet
                Worksheet activeSheet = workbook.Worksheets[0];

                // Add sample data
                activeSheet.Cells["A1"].PutValue("Name");
                activeSheet.Cells["B1"].PutValue("Age");
                activeSheet.Cells["A2"].PutValue("John");
                activeSheet.Cells["B2"].PutValue(30);
                activeSheet.Cells["A3"].PutValue("Alice");
                activeSheet.Cells["B3"].PutValue(25);

                // Ensure the first worksheet is the active one (optional)
                workbook.Worksheets.ActiveSheetIndex = 0;

                // Configure CSV save options with semicolon separator
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ';',
                    ExportAllSheets = false
                };

                // Save the active worksheet to a CSV file
                workbook.Save("ActiveSheetExport.csv", csvOptions);

                Console.WriteLine("Active worksheet exported to CSV with semicolon separator.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportActiveSheetToCsv.Run();
        }
    }
}
