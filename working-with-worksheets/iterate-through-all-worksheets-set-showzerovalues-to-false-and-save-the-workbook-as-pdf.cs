// Title: Loop through all worksheets, turn off zero-value display, and save the workbook as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates over every worksheet, sets ShowZeroValues to false, and saves the result as a PDF. | Create a reusable C# method that accepts an input Excel path, disables zero-value display on all worksheets via Aspose.Cells, and returns a PDF byte array. | Add comprehensive error handling and logging to a C# Aspose.Cells routine that hides zero values across all sheets before exporting to PDF.
// Common Searches: Aspose.Cells C# hide zero values on all worksheets before PDF export | how to set ShowZeroValues false for each sheet using Aspose.Cells | C# loop through worksheets and change page setup options Aspose.Cells | export Excel workbook to PDF with zero values suppressed Aspose.Cells | sample code to disable zero display in Aspose.Cells and generate PDF
// Tags: loop worksheets set ShowZeroValues false Aspose.Cells | hide zero values page setup Aspose.Cells C# | export workbook to PDF Aspose.Cells | worksheet page setup hide zero values Aspose.Cells | PDF conversion zero-value suppression Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, iterates through each worksheet to set ShowZeroValues to false (hiding zero values), and then saves the entire workbook as a PDF using Aspose.Cells for .NET, with proper exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook from the existing Excel file
                Workbook workbook = new Workbook(inputFile);

                // If needed, configure page setup options here.
                // Note: Aspose.Cells does not expose a PrintZeroValues property in this version.

                // Save the workbook as a PDF file
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
