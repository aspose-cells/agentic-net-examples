// Title: Convert a macro‑enabled XLSM workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: how to export an XLSM file to PDF using Aspose.Cells in C# | Aspose.Cells C# convert macro enabled Excel workbook to PDF | C# code example for saving an XLSM workbook as PDF with Aspose.Cells | save Excel macro workbook as PDF using Aspose.Cells .NET library | load XLSM workbook and generate PDF in C# with Aspose.Cells
// Tags: Aspose.Cells XLSM to PDF conversion C# | export macro-enabled Excel to PDF Aspose.Cells | C# workbook.Save PDF Aspose.Cells example | load XLSM workbook Aspose.Cells .NET | Aspose.Cells PDF export without macros

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The sample program verifies that an input.xlsm file exists, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF file (output.pdf) using the SaveFormat.Pdf enumeration. It includes basic error handling to report any issues during the conversion.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsm";
                string outputFile = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the XLSM workbook
                Workbook workbook = new Workbook(inputFile);

                // Save the workbook as PDF
                workbook.Save(outputFile, SaveFormat.Pdf);

                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
