// Title: Export a Selected Cell Range to CSV with Proper Comma Escaping using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill cells with text that contains commas, define a specific range (A1:B3), and save that range as a UTF‑8 CSV file. The example configures TxtSaveOptions with a comma separator and QuoteType.Normal so any value that includes a comma is automatically quoted, guaranteeing correct CSV formatting.
// Keywords: Aspose.Cells CSV export | export range to CSV C# | quote commas Aspose.Cells | TxtSaveOptions QuoteType.Normal | selected cells CSV .NET | escape delimiter in CSV | Aspose.Cells range export | C# CSV generation Aspose
// Common Searches: Aspose.Cells export specific range to CSV | how to quote fields with commas in Aspose.Cells | save worksheet area as CSV using TxtSaveOptions | C# export A1:B3 to CSV with proper escaping | Aspose.Cells CSV delimiter handling
// Developer Intent: Save a defined cell block as a CSV file while automatically quoting any cell text that contains the delimiter.
// Use Cases: Produce a CSV report from a subset of a workbook where description columns may include commas. | Integrate Excel data with external systems that require correctly quoted CSV files. | Export filtered or calculated ranges without writing the entire worksheet to disk.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a chosen range to CSV, ensuring commas inside values are quoted. | Explain the effect of TxtSaveOptions.QuoteType.Normal on CSV output in Aspose.Cells. | Adapt the sample to use a semicolon as the delimiter while still quoting fields that contain the delimiter.

using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCsvExport
{
    // Demonstrates how to create a workbook, fill cells with text that contains commas, define a specific range (A1:B3), and save that range as a UTF‑8 CSV file. The example configures TxtSaveOptions with a comma separator and QuoteType.Normal so any value that includes a comma is automatically quoted, guaranteeing correct CSV formatting.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data, including commas inside text
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Description");
                cells["A2"].PutValue("Item1");
                cells["B2"].PutValue("Size, Large");          // comma inside text
                cells["A3"].PutValue("Item2");
                cells["B3"].PutValue("Color, Red, Blue");    // multiple commas

                // Define the range to export (A1:B3)
                AsposeRange exportRange = cells.CreateRange("A1:B3");

                // Build the CellArea representing the range
                CellArea area = new CellArea
                {
                    StartRow = exportRange.FirstRow,
                    StartColumn = exportRange.FirstColumn,
                    EndRow = exportRange.FirstRow + exportRange.RowCount - 1,
                    EndColumn = exportRange.FirstColumn + exportRange.ColumnCount - 1
                };

                // Configure text save options for CSV
                TxtSaveOptions saveOptions = new TxtSaveOptions
                {
                    Separator = ',',                                 // Use comma as delimiter
                    QuoteType = TxtValueQuoteType.Normal,            // Quote values that contain commas
                    ExportArea = area,                               // Export only the defined range
                    TrimLeadingBlankRowAndColumn = true,
                    Encoding = Encoding.UTF8
                };

                // Ensure the output directory exists
                string outputPath = "ExportedData.csv";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the selected range as a CSV file; commas inside text will be quoted
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"CSV file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
