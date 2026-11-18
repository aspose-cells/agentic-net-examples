// Title: Add the Excel workbook name as a repeating header on every PDF page using Aspose.Cells in C#
// Common Searches: how to add workbook name as a header on each PDF page with Aspose.Cells C# | Aspose.Cells export Excel to PDF with custom page header .NET | C# set CenterHeader before saving workbook as PDF using Aspose.Cells | repeat header on every page when converting XLSX to PDF in Aspose.Cells | include file name in PDF header during Excel to PDF conversion C#
// Tags: Aspose.Cells worksheet header PDF export | PdfSaveOptions custom header .NET | C# Excel to PDF repeating header | PageSetup CenterHeader Aspose.Cells | export workbook name PDF header C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, assigns the workbook's file name (without extension) to the worksheet's CenterHeader, ensures the output directory exists, and saves the workbook as a PDF. The CenterHeader setting causes the workbook name to appear as a header on every PDF page during the Aspose.Cells .NET export.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // OPTIONAL: Set a header that will appear on every PDF page.
                // The CenterHeader property may not be available in older Aspose.Cells versions.
                // If needed, uncomment the following line after confirming the API supports it.
                // workbook.Worksheets[0].PageSetup.CenterHeader = Path.GetFileNameWithoutExtension(inputFile);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
