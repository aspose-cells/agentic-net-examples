// Title: Define a worksheet print area and save as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to assign a cell range as the worksheet's print area and then generate a PDF file. | Show how to configure portrait orientation, fit‑to‑page dimensions, and create the target folder before saving a workbook as PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet print area for PDF export | define print area without named range Aspose.Cells .NET | export specific range to PDF using Aspose.Cells in C# | page setup fit to page and orientation Aspose.Cells PDF | create output directory automatically when saving PDF Aspose.Cells
// Tags: Aspose.Cells set print region C# | Aspose.Cells export PDF from worksheet | Aspose.Cells page setup fit to page | Aspose.Cells portrait orientation PDF | Aspose.Cells create output folder

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, fills cells A1‑B2, assigns the print area via PageSetup.PrintArea, sets portrait orientation and fit‑to‑page options, ensures the destination folder exists, and saves the worksheet as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Define the print area directly (no need for a named range)
            string printAreaAddress = "'Sheet1'!$A$1:$B$2";
            sheet.PageSetup.PrintArea = printAreaAddress;

            // Optional page setup adjustments
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 1;

            // Define output file path
            string outputPath = "PrintedSheet.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
