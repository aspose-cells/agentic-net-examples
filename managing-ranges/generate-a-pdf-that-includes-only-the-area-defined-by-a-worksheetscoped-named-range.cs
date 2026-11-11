// Title: Create a PDF containing only a worksheet‑scoped named range with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export only a specific worksheet‑scoped named range to PDF | How to set print area from a worksheet‑scoped named range before saving as PDF using Aspose.Cells | Aspose.Cells .NET generate PDF from Excel named range without exporting the whole sheet | Save selected range as PDF in C# with Aspose.Cells workbook.Save | Retrieve worksheet‑scoped named range and convert to PDF using Aspose.Cells API
// Tags: worksheet scoped named range print area Aspose.Cells | export selected Excel range to PDF C# | Aspose.Cells set print area from named range | save workbook as PDF using print area .NET | C# generate PDF from Excel named range

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, obtains a worksheet‑scoped named range, assigns its address to the worksheet's print area, and saves the workbook as a PDF so that only the defined range is exported.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";
                const string sheetName = "Sheet1";
                const string rangeName = "MyRange";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Get the target worksheet
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return;
                }

                // Retrieve the named range (worksheet‑scoped) from the workbook's Names collection
                Name scopedRange = workbook.Worksheets.Names[rangeName];
                if (scopedRange == null)
                {
                    Console.WriteLine($"Named range '{rangeName}' not found.");
                    return;
                }

                // Set the print area to the range address (remove leading '=' if present)
                string area = scopedRange.RefersTo;
                if (area.StartsWith("="))
                    area = area.Substring(1);
                sheet.PageSetup.PrintArea = area;

                // Save the workbook as PDF (only the print area will be exported)
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine("PDF generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
