// Title: Set worksheet print scaling to fit all columns on one page using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, set the first worksheet's PageSetup.FitToPagesWide to 1 and FitToPagesTall to 0, then save the file. | Write C# code that adjusts the print layout of an existing .xlsx so that columns fit within a single page width while rows may continue onto additional pages, using Aspose.Cells. | Create a .NET program that opens an Excel file, configures the worksheet's print scaling to one-page width without distortion, and outputs the modified workbook.
// Common Searches: Aspose.Cells C# set worksheet to print all columns on one page width | How to configure FitToPagesWide in Aspose.Cells .NET | Print scaling Excel worksheet to fit columns using Aspose.Cells library | C# Aspose.Cells fit columns to single printed page without affecting rows | Adjust page setup for Excel print layout programmatically with Aspose.Cells
// Tags: Aspose.Cells column scaling configuration | worksheet page configuration .NET | single-page width column fit C# | Excel workbook scaling preserving layout | programmatic page scaling with Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, sets the first worksheet's PageSetup to fit all columns on a single printed page width (FitToPagesWide = 1, FitToPagesTall = 0), and saves the updated file.
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

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Fit all columns on one page width; rows can span multiple pages
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
