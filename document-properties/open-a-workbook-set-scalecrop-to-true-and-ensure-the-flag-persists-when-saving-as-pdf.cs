// Title: Fit every worksheet to a single page and convert an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file using Aspose.Cells, set each worksheet's PageSetup.FitToPagesWide and FitToPagesTall to 1, then save the workbook as a PDF. | Programmatically enforce a one‑page‑wide and one‑page‑tall layout for all sheets before exporting the workbook to PDF in C# with Aspose.Cells.
// Common Searches: Aspose.Cells C# fit all worksheets to one page before PDF export | how to set FitToPagesWide and FitToPagesTall for each sheet in Aspose.Cells | preserve page scaling when saving Excel workbook as PDF using Aspose.Cells .NET | C# code to force single-page layout for Excel sheets in PDF conversion | Aspose.Cells page setup scaling options for PDF output
// Tags: set FitToPagesWide FitToPagesTall Aspose.Cells | scale worksheets to single page PDF conversion .NET | page setup scaling for PDF export Aspose.Cells | apply uniform page scaling across workbook C# | export Excel workbook to PDF with page fit Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, iterates through all worksheets to set PageSetup.FitToPagesWide and FitToPagesTall to 1 (forcing each sheet onto a single page), and then saves the workbook as a PDF, preserving the scaling settings in the output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure each worksheet to scale the content to fit a single page
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Fit the worksheet to one page wide and one page tall
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;
            }

            // Save the workbook as PDF; the scaling settings are applied in the PDF output
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
