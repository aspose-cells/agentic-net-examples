// Title: Set FitToPagesTall = 1 while keeping FitToPagesWide default and export a multi‑sheet PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Configure each worksheet's PageSetup to set FitToPagesTall = 1, leave FitToPagesWide unchanged, and save the workbook as a multi‑sheet PDF using Aspose.Cells in C#. | Create a PDF where every sheet fits vertically on a single page while width scaling stays automatic by applying per‑sheet FitToPagesTall and using PdfSaveOptions with AllColumnsInOnePagePerSheet set to false.
// Common Searches: Aspose.Cells C# set FitToPagesTall to 1 without changing FitToPagesWide for PDF export | How to export each Excel worksheet to a single‑page‑height PDF while preserving automatic width scaling | C# Aspose.Cells multi‑sheet PDF with vertical fit only per worksheet | PdfSaveOptions settings to respect worksheet FitToPagesTall in Aspose.Cells | FitToPagesTall 1 PDF conversion Aspose.Cells .NET example
// Tags: page-setup vertical page fit per worksheet | maintain automatic width scaling during pdf export | multi-sheet pdf generation with per-sheet scaling | c# pdfsaveoptions allcolumnsononepagepersheet false | aspocells page-setup scaling example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads input.xlsx, sets each worksheet's PageSetup.FitToPagesTall to 1 while leaving FitToPagesWide at its default, configures PdfSaveOptions to honor those per‑sheet settings, and saves the workbook as a multi‑sheet PDF (output.pdf) with consistent vertical scaling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply FitToPagesTall = 1 to every worksheet while leaving FitToPagesWide unchanged
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.FitToPagesTall = 1; // Force each sheet to fit into one page vertically
                // pageSetup.FitToPagesWide remains at its default (0) – automatic width scaling
            }

            // Configure PDF save options to preserve the per‑sheet scaling settings
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = false, // Respect each sheet's width setting
                OnePagePerSheet = false               // Allow multiple pages per sheet if needed
            };

            // Save the workbook as a multi‑sheet PDF with consistent scaling
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
