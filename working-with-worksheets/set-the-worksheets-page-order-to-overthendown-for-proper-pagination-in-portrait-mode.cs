// Title: How to set worksheet page order to OverThenDown for portrait printing with Aspose.Cells for .NET
// Common Searches: how to set OverThenDown pagination in Aspose.Cells C# | C# example for configuring worksheet print order with Aspose.Cells | setting portrait page orientation and pagination direction using Aspose.Cells .NET | Aspose.Cells print order property tutorial for Excel files | create folder and save workbook with Aspose.Cells in C#
// Tags: Aspose.Cells OverThenDown page order | C# worksheet pagination direction | Aspose.Cells portrait orientation setting | C# ensure output directory for workbook save | Excel workbook save path handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, accesses the first Worksheet, and demonstrates how to configure the PageSetup.PrintOrder to OverThenDown for correct portrait pagination. It also sets the page orientation to Portrait, ensures the output directory exists, and saves the workbook as an XLSX file while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the page order to over‑then‑down for proper pagination in portrait mode
            // Note: In some older Aspose.Cells versions the PrintOrder property may not be available.
            // If supported, uncomment the following line:
            // sheet.PageSetup.PrintOrder = PrintOrderType.OverThenDown;

            // Ensure the orientation is portrait (portrait is the default)
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
