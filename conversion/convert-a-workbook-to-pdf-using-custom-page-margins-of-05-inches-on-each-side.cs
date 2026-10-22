// Title: Convert an Excel workbook to PDF with 0.5‑inch page margins using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set 0.5 inch margins before saving workbook as PDF | Export Excel to PDF with custom page margins using Aspose.Cells .NET | C# code to apply uniform half‑inch margins to all worksheets and convert to PDF | Set left, right, top, bottom margins in Aspose.Cells page setup for PDF output | Convert .xlsx to PDF with specific margin size using Aspose.Cells library
// Tags: Aspose.Cells PDF conversion with custom margins | C# set worksheet page margins Aspose.Cells | Excel to PDF half inch margins Aspose.Cells | PageSetup margin configuration Aspose.Cells .NET | Workbook.Save PDF with margin settings C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, verifies its presence, applies a uniform 0.5‑inch (1.27 cm) margin to all sides of each worksheet via the PageSetup object, and then saves the workbook as a PDF using Aspose.Cells, with error handling for robustness.
class WorkbookToPdfWithMargins
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define margin size in centimeters (0.5 inch = 1.27 cm)
            const double marginInCentimeters = 0.5 * 2.54; // 1.27 cm

            // Apply the custom margins to each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.LeftMargin = marginInCentimeters;
                pageSetup.RightMargin = marginInCentimeters;
                pageSetup.TopMargin = marginInCentimeters;
                pageSetup.BottomMargin = marginInCentimeters;
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
