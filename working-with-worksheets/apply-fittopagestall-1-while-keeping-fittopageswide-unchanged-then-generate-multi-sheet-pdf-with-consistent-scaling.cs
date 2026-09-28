// Title: Apply FitToPagesTall = 1 to all worksheets while keeping existing FitToPagesWide and export as a multi‑sheet PDF with Aspose.Cells for .NET
// AI Prompts: Iterate through every worksheet, set the page setup's FitToPagesTall property to 1 without altering the width scaling, then save the workbook as a PDF. | Create a multi‑sheet PDF where each sheet is constrained to one page tall while preserving its original width scaling using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# set each worksheet to fit one page tall keep current width scaling | export Excel workbook to PDF with uniform page‑height scaling across all sheets Aspose.Cells | retain original width scaling while setting page height to one page for all worksheets in .NET
// Tags: fit page tall scaling Aspose.Cells | maintain existing width scaling Aspose.Cells | multi‑sheet PDF export C# Aspose.Cells | apply page setup to all worksheets Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel file, forces every worksheet to fit to a single page tall while preserving its existing width scaling, and saves the workbook as a multi‑sheet PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply FitToPagesTall = 1 to every worksheet while preserving FitToPagesWide
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;

                // Preserve the existing FitToPagesWide value (no change needed)
                int currentFitWide = pageSetup.FitToPagesWide;

                // Set FitToPagesTall to 1
                pageSetup.FitToPagesTall = 1;

                // Reassign FitToPagesWide to keep the original setting (optional)
                pageSetup.FitToPagesWide = currentFitWide;
            }

            // Save the workbook as a multi‑sheet PDF with the applied scaling
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
