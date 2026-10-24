// Title: Clear existing printer settings and apply new print area, orientation, and margins to an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# reset worksheet printer settings before setting print area | how to clear page setup in an Excel worksheet with Aspose.Cells .NET | remove previous print configuration from a worksheet using Aspose.Cells C# example | Aspose.Cells clear existing margins and orientation then apply new print settings | reset printer options on Excel sheet before configuring print layout with Aspose.Cells
// Tags: clear worksheet page setup Aspose.Cells | reset printer settings Excel .NET | set print area landscape orientation Aspose.Cells | configure print margins Aspose.Cells C# | modify worksheet print configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, accesses the first worksheet, overwrites any prior printer settings by configuring the PageSetup (print area, landscape orientation, and custom margins), and saves the modified file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = @"C:\Input\Sample.xlsx";
            string outputPath = @"C:\Output\Sample_Modified.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Configure print settings
                var pageSetup = worksheet.PageSetup;
                pageSetup.PrintArea = "A1:D20";
                pageSetup.Orientation = PageOrientationType.Landscape;
                pageSetup.LeftMargin = 0.5;   // inches
                pageSetup.RightMargin = 0.5;  // inches
                pageSetup.TopMargin = 0.75;   // inches
                pageSetup.BottomMargin = 0.75; // inches

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
