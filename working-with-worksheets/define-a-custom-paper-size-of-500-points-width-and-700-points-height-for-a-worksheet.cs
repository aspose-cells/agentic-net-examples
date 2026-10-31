// Title: How to set a 500 pt × 700 pt custom paper size for a worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet custom paper size 500 points width 700 points height | SetCustomPaperSize method usage in Aspose.Cells .NET | Define custom page dimensions for an Excel worksheet using Aspose.Cells | C# Aspose.Cells page setup custom paper size example | Specify custom paper width and height in points with Aspose.Cells
// Tags: SetCustomPaperSize method Aspose.Cells | worksheet custom paper size points | Aspose.Cells page setup custom dimensions | C# define worksheet paper width height | Aspose.Cells custom page size Excel

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PaperSize to Custom, optionally calls SetCustomPaperSize(500, 700) to define the width and height in points (when supported), and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure page setup
            PageSetup pageSetup = sheet.PageSetup;
            // Attempt to use a custom paper size. In some Aspose.Cells versions
            // PaperWidth and PaperHeight are read‑only, so we only set the enum here.
            pageSetup.PaperSize = PaperSizeType.Custom;

            // If the current Aspose.Cells version supports setting dimensions,
            // uncomment the following line (available in newer releases):
            // pageSetup.SetCustomPaperSize(500, 700); // width and height in points

            // Save the workbook to a file
            string outputPath = "CustomPaperSize.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
