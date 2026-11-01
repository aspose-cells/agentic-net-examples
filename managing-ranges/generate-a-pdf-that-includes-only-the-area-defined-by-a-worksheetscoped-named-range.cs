// Title: Create a PDF that contains only a worksheet‑scoped named range using Aspose.Cells for .NET (C#)
// Common Searches: how to export only a worksheet scoped named range to PDF with Aspose.Cells in C# | Aspose.Cells C# set print area from a named range before saving as PDF | C# generate PDF from a specific Excel named range using Aspose.Cells | save Excel worksheet as PDF limited to a named range with Aspose.Cells | Aspose.Cells .NET define print area using worksheet scoped name for PDF output
// Tags: Aspose.Cells set print area from worksheet scoped name | C# PDF export of Excel named range using Aspose.Cells | define print area via named range Aspose.Cells .NET | save workbook as PDF limited to specific range C# | Aspose.Cells worksheet scoped named range PDF generation

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, retrieves a worksheet‑scoped named range, assigns its reference to the worksheet's print area, and saves the workbook as a PDF so that only the defined range appears in the output, with basic file checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";
            const string rangeName = "MyRange";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range from the workbook's Names collection
            // This collection includes both workbook‑scoped and worksheet‑scoped names
            Name scopedRange = workbook.Worksheets.Names[rangeName];

            if (scopedRange == null)
            {
                Console.WriteLine($"The specified named range \"{rangeName}\" was not found.");
                return;
            }

            // Set the worksheet's print area to the range referenced by the named range
            sheet.PageSetup.PrintArea = scopedRange.RefersTo;

            // Save the workbook as PDF; only the defined print area will be included
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
