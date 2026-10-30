// Title: Export an Aspose.Cells workbook to CSV in C# while preserving blank rows using TxtSaveOptions
// Common Searches: C# Aspose.Cells export workbook to CSV keep blank rows | How to retain empty lines when saving Excel as CSV with Aspose.Cells | Using TxtSaveOptions to preserve blank rows in CSV output in C# | Aspose.Cells CSV save options for custom separator and empty rows | Exporting Excel to CSV with Aspose.Cells without removing blank rows
// Tags: Aspose.Cells TxtSaveOptions CSV export | preserve empty rows CSV Aspose.Cells | custom CSV separator Aspose.Cells C# | export workbook to CSV with blank lines | C# Aspose.Cells save workbook as CSV

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, adds header and data rows with an intentional blank row, configures TxtSaveOptions for CSV format using a comma separator, ensures the output directory exists, saves the workbook to 'output.csv', and logs the file path while handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data with a blank row between rows 2 and 4
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Row1Col1");
                sheet.Cells["B2"].PutValue("Row1Col2");
                // Row 3 is intentionally left blank
                sheet.Cells["A4"].PutValue("Row2Col1");
                sheet.Cells["B4"].PutValue("Row2Col2");

                // Configure CSV save options (separator = comma)
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ',' // Use comma as delimiter
                };

                // Define output file path and ensure directory exists
                string outputPath = "output.csv";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as CSV using the custom options
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
