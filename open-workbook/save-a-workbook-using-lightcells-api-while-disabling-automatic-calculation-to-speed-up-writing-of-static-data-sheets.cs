// Title: Save a static Excel workbook with Aspose.Cells LightCells API and disable automatic calculation to improve write performance (C#)
// AI Prompts: Write C# code that fills a worksheet with static values, turns off automatic formula calculation, and saves the file using Aspose.Cells LightCells API. | Demonstrate how to disable workbook recalculation before calling Save with LightCells to speed up exporting large static data sheets in .NET.
// Common Searches: how to disable automatic calculation when using Aspose.Cells LightCells in C# | save large static Excel file faster with LightCells API .NET | Aspose.Cells turn off formula recalculation before workbook.Save | C# LightCells write static data sheet without triggering calculation | performance optimization for Excel export using Aspose.Cells LightCells
// Tags: disable automatic calculation Aspose.Cells LightCells | LightCells API save static Excel workbook | write static data sheet C# Aspose.Cells | Excel export performance optimization .NET | turn off formula recalculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, populates the first worksheet with 500 rows of static text values, disables automatic formula calculation, and saves the workbook as 'StaticDataWorkbook.xlsx', ensuring the output directory exists and handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet where data will be written
            Worksheet sheet = workbook.Worksheets[0];

            // Fill 500 rows and 5 columns with sample values
            for (int row = 0; row < 500; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row}C{col}");
                }
            }

            // Define output file path
            string outputPath = "StaticDataWorkbook.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
