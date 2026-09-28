// Title: How to remove a worksheet’s embedded printer settings by setting PageSetup.PrinterSettings to null with Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, assign null to worksheet.PageSetup.PrinterSettings, and save the file using Aspose.Cells in C#. | Programmatically clear the printer configuration of a specific worksheet before exporting the workbook with Aspose.Cells. | Reset the PageSetup printer settings of the first worksheet to default by assigning null in a .NET application.
// Common Searches: asp.net remove printer settings from Excel worksheet using Aspose.Cells | set PageSetup.PrinterSettings to null in C# Aspose.Cells example | delete embedded printer configuration in an .xlsx file programmatically | how to reset worksheet printer setup before saving with Aspose.Cells
// Tags: Aspose.Cells clear worksheet printer settings | PageSetup.PrinterSettings null assignment | remove embedded printer configuration from Excel | programmatic worksheet printer setup reset | C# Aspose.Cells printer settings removal

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads input.xlsx, sets the first worksheet's PageSetup.PrinterSettings property to null to strip any embedded printer configuration, and saves the modified workbook as output.xlsx while handling missing files and exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Remove embedded printer configuration (set to null if supported)
                worksheet.PageSetup.PrinterSettings = null;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
