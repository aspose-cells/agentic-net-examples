// Title: How to set a worksheet's print area to the named range "ReportArea" using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing Excel file with Aspose.Cells, verifies the file exists, assigns the named range "ReportArea" to the worksheet's PageSetup.PrintArea, and saves the workbook. | Create a C# snippet that accesses the first worksheet in a workbook, sets its print area to a predefined named range, and includes proper exception handling for missing files.
// Common Searches: Aspose.Cells C# set worksheet print area to a named range called ReportArea | How to use PageSetup.PrintArea with a named range in Aspose.Cells .NET | C# code for dynamic printing area based on a named range in Excel using Aspose.Cells
// Tags: Aspose.Cells PageSetup.PrintArea named range | C# set Excel print area Aspose.Cells | dynamic print area Excel Aspose.Cells .NET | load workbook assign named range print area | handle missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies that an input Excel file exists, loads it with Aspose.Cells, accesses the first worksheet, sets its PageSetup.PrintArea to the named range "ReportArea", saves the modified workbook to a new file, and reports any errors.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or specify by name if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Set the print area to the named range "ReportArea"
                sheet.PageSetup.PrintArea = "ReportArea";

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
