// Title: Unhide column B, set its width to 50 characters, and export the worksheet to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to unhide a specific column and define its width before saving as PDF | set column B width to 50 characters using Aspose.Cells .NET | export Excel worksheet to PDF while preserving column visibility with Aspose.Cells | Aspose.Cells hide/unhide columns programmatically in C# | adjust column width in Aspose.Cells and generate PDF output
// Tags: unhide column Aspose.Cells C# | set column width Aspose.Cells | column visibility PDF export Aspose.Cells | Aspose.Cells worksheet to PDF conversion | adjust column formatting before PDF save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, unhides column B, sets its width to 50 characters, and saves the worksheet as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unhide column B (zero‑based index 1) and set its width
                sheet.Cells.Columns[1].IsHidden = false;
                // Width is specified in characters; 50 characters approximates the desired size
                sheet.Cells.SetColumnWidth(1, 50);

                // Save the workbook as PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
