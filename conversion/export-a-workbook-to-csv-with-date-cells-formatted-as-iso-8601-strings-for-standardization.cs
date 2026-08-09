// Title: Export a Workbook to CSV with ISO‑8601 Date Formatting using Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create an Aspose.Cells workbook, apply the ISO‑8601 pattern (yyyy‑MM‑ddTHH:mm:ss) to DateTime cells, ensure the target folder exists, and save the file as CSV. The resulting CSV contains standardized date strings suitable for data exchange and logging.
// Keywords: aspose.cells | csv export | iso 8601 | c# | dotnet | custom date format | saveformat.csv | excel to csv conversion | date formatting in csv | data integration
// Common Searches: Aspose.Cells export workbook to CSV with ISO 8601 dates | C# set custom date format before CSV save Aspose | How to format DateTime cells as yyyy-MM-ddTHH:mm:ss in CSV using Aspose.Cells | Create CSV file with standardized timestamps in .NET | Aspose.Cells CSV export timezone handling
// Developer Intent: Generate a CSV file from a workbook where every DateTime value is written in ISO‑8601 format.
// Use Cases: Exchange CSV data with APIs or services that require ISO‑8601 timestamps. | Produce audit‑ready reports with consistent date strings across platforms. | Load CSV into databases that expect the "yyyy-MM-ddTHH:mm:ss" pattern.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to CSV with dates formatted as "yyyy-MM-ddTHH:mm:ss" and creates the output directory if missing. | Explain how to apply a custom style to an entire column of DateTime cells before saving as CSV with Aspose.Cells. | Show how to include UTC offset or 'Z' designator when exporting dates to CSV in ISO‑8601 using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // This C# example demonstrates how to create an Aspose.Cells workbook, apply the ISO‑8601 pattern (yyyy‑MM‑ddTHH:mm:ss) to DateTime cells, ensure the target folder exists, and save the file as CSV. The resulting CSV contains standardized date strings suitable for data exchange and logging.
    public class ExportWorkbookToCsvWithIsoDates
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("CreatedOn");

                cells["A2"].PutValue(1);
                cells["B2"].PutValue("Alice");
                cells["C2"].PutValue(new DateTime(2023, 5, 15, 14, 30, 0)); // sample date

                cells["A3"].PutValue(2);
                cells["B3"].PutValue("Bob");
                cells["C3"].PutValue(DateTime.Now); // current date‑time

                // Define ISO 8601 format for date cells
                const string isoFormat = "yyyy-MM-ddTHH:mm:ss";

                // Apply the custom format to all cells in column C that contain dates
                for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
                {
                    Cell dateCell = cells[row, 2]; // column index 2 = "C"
                    if (dateCell.Type == CellValueType.IsDateTime)
                    {
                        Style style = dateCell.GetStyle();
                        style.Custom = isoFormat; // set ISO format
                        dateCell.SetStyle(style);
                    }
                }

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ExportedData.csv");

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as CSV (lifecycle rule: save)
                // The custom style ensures dates are written in ISO 8601 format.
                workbook.Save(outputFile, SaveFormat.Csv);

                Console.WriteLine($"Workbook exported to CSV with ISO 8601 date formatting: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
