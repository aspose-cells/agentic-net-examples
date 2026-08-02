// Title: Enable Black‑and‑White Printing for All Worksheets in Aspose.Cells (C#)
// Description: Creates a new Workbook, loops through each Worksheet, sets PageSetup.BlackAndWhite to true to force grayscale printing, and saves the file. Shows how to reduce ink consumption across an entire Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells black and white printing | C# grayscale worksheet printing | PageSetup.BlackAndWhite property | reduce ink usage Excel .NET | print workbook in grayscale | Aspose.Cells printing settings | disable color printing Aspose
// Common Searches: Aspose.Cells enable black and white printing | C# set workbook to grayscale printing | turn off color printing in Excel using Aspose.Cells | reduce ink consumption Aspose.Cells .NET | global page setup black and white Aspose
// Developer Intent: Apply grayscale (black‑and‑white) printing to every worksheet in a workbook via Aspose.Cells for .NET.
// Use Cases: Generate cost‑effective reports that print in grayscale to save ink. | Create archival copies that comply with corporate policies requiring monochrome prints. | Prepare bulk‑exported Excel files for mass printing where color adds no value.
// AI Prompts: Show how to enable black‑and‑white printing for a single worksheet using Aspose.Cells in C#. | Provide code to toggle grayscale printing based on a configuration flag in a .NET application. | Explain how to revert a workbook back to color printing after it was set to black‑and‑white.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, loops through each Worksheet, sets PageSetup.BlackAndWhite to true to force grayscale printing, and saves the file. Shows how to reduce ink consumption across an entire Excel workbook using Aspose.Cells for .NET.
    public class SetBlackAndWhitePrinting
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable black‑and‑white printing for every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // PageSetup.BlackAndWhite property turns on B&W printing
                    sheet.PageSetup.BlackAndWhite = true;
                }

                // Save the workbook with the updated settings
                string outputPath = "BlackAndWhiteWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetBlackAndWhitePrinting.Run();
        }
    }
}
