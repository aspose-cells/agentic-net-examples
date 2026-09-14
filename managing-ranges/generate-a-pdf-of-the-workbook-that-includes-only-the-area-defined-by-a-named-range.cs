// Title: Generate a PDF containing only a named range from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create C# code that loads an Excel file, locates a named range, sets the worksheet's print area to that range, and saves the selected area as a PDF with Aspose.Cells. | Show how to extract the worksheet name from a named range's RefersTo string, adjust page‑setup options, and export only that range to PDF using Aspose.Cells in C#. | Provide a step‑by‑step example that verifies the input file, creates missing output directories, and handles errors while converting a named range to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# export only a named range to PDF | set print area from named range before saving workbook as PDF using Aspose.Cells | retrieve worksheet name from named range RefersTo string Aspose.Cells | C# generate PDF of specific Excel range with page setup using Aspose.Cells
// Tags: named range PDF export Aspose.Cells | set worksheet print area programmatically C# | retrieve named range address Aspose.Cells | configure page setup for PDF Aspose.Cells | handle file existence and output directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, finds the named range "MyRange", determines its worksheet, sets the worksheet's print area to the range address, optionally configures page‑setup settings, and saves the workbook as a PDF that contains only the defined area.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange == null)
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                    return;
                }

                // Get the actual range object that the name refers to
                AsposeRange range = namedRange.GetRange();

                // Determine the worksheet that contains the named range
                // The RefersTo property contains something like "Sheet1!$A$1:$B$2"
                string refersTo = namedRange.RefersTo;
                string sheetName = refersTo.Split('!')[0].Trim('\'');
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return;
                }

                // Set the print area of the worksheet to the address of the named range
                sheet.PageSetup.PrintArea = range.Address;

                // Optional page setup adjustments
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 0; // 0 means as many pages tall as needed

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF; only the defined print area will be exported
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
