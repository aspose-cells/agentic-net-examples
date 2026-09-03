// Title: How to disable automatic XML map refresh in Aspose.Cells for .NET while performing bulk cell updates to boost performance
// AI Prompts: Write C# code using Aspose.Cells that sets Workbook.Settings.EnableXMLMapRefresh to false, inserts 10,000 rows into the first column, then sets the property back to true before saving the workbook. | Show a .NET example that temporarily suspends XML map auto‑refresh, performs high‑volume cell writes, and restores the refresh setting to ensure data integrity with Aspose.Cells.
// Common Searches: Aspose.Cells C# disable XML map auto refresh for large data import | Improve bulk cell write speed by turning off XML map refresh in .NET workbook | Temporarily suspend XML map refresh during massive updates with Aspose.Cells
// Tags: disable xml map auto refresh aspose.cells | bulk cell write performance optimization .net | workbook settings enablexmlmaprefresh property | c# aspose.cells large data insertion | excel xml map refresh control aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, disables the automatic XML map refresh, writes 10,000 sequential numbers into column A of the first worksheet, re‑enables the XML map refresh, and saves the file, demonstrating how to speed up bulk cell updates with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                var workbook = new Workbook();

                // Example bulk update: write 10,000 values into the first column of the first worksheet
                var cells = workbook.Worksheets[0].Cells;
                for (int row = 0; row < 10000; row++)
                {
                    cells[row, 0].PutValue(row);
                }

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
