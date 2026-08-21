// Title: Export a Worksheet Range to CSV with Proper Comma Escaping – Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills cells with values containing commas, defines a range (A1:B3) and saves only that range to a CSV file. TxtSaveOptions are configured with a comma separator and QuoteType.Normal so any commas inside cell text are automatically quoted, producing a valid CSV.
// Keywords: Aspose.Cells CSV export | export range to CSV .NET | comma escaping in CSV | TxtSaveOptions QuoteType.Normal | ExportArea cell range | C# Aspose.Cells example | save selected cells as CSV
// Common Searches: how to export a selected range to CSV using Aspose.Cells | Aspose.Cells CSV export with commas inside text | C# save worksheet area as CSV file | TxtSaveOptions quote values containing commas | export A1:B3 range to CSV Aspose.Cells
// Developer Intent: Save a specific cell range as a CSV file while automatically quoting values that contain commas.
// Use Cases: Generate CSV reports where description fields include commas without breaking column alignment. | Provide data extracts for APIs or downstream systems that require correctly escaped CSV values. | Automate the creation of email attachments containing only a subset of worksheet data.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet range to CSV and ensure commas are escaped. | Explain how TxtSaveOptions.QuoteType.Normal quotes cell values that contain delimiters during CSV export. | Show how to configure ExportArea in TxtSaveOptions to save only the range A1:B3 as a CSV file.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// C# example that creates a workbook, fills cells with values containing commas, defines a range (A1:B3) and saves only that range to a CSV file. TxtSaveOptions are configured with a comma separator and QuoteType.Normal so any commas inside cell text are automatically quoted, producing a valid CSV.
class ExportRangeToCsv
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with data, including commas inside text
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Description");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue("Size, Large");
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue("Color, Red");

            // Define the range to be exported (A1:B3)
            AsposeRange exportRange = worksheet.Cells.CreateRange("A1:B3");

            // Configure text save options for CSV export
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',', // comma delimiter
                QuoteType = TxtValueQuoteType.Normal, // quote values that contain commas
                ExportArea = new CellArea
                {
                    StartRow = exportRange.FirstRow,
                    EndRow = exportRange.FirstRow + exportRange.RowCount - 1,
                    StartColumn = exportRange.FirstColumn,
                    EndColumn = exportRange.FirstColumn + exportRange.ColumnCount - 1
                }
            };

            // Determine output file path and ensure directory exists
            string outputPath = "ExportedRange.csv";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the selected range to a CSV file; commas inside text will be properly escaped
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Range exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
