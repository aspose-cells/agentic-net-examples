// Title: Convert an Excel workbook to PDF without scaling using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export Excel to PDF preserving original page size | disable automatic scaling when saving workbook as PDF with Aspose.Cells | set worksheet page setup zoom 100% for PDF conversion Aspose.Cells .NET | prevent FitToPagesWide and FitToPagesTall from affecting PDF output in Aspose.Cells | convert XLSX to PDF without page fit scaling using Aspose.Cells in C#
// Tags: Aspose.Cells PDF export without scaling | page setup scaling options Aspose.Cells | disable FitToPagesWide FitToPagesTall Aspose.Cells | preserve original dimensions Excel to PDF C# | SaveFormat.Pdf scaling control Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program checks for the input XLSX file, loads it with Aspose.Cells (handling password errors), iterates through each worksheet to set PageSetup.Zoom to 100 and both FitToPagesWide and FitToPagesTall to 0, thereby disabling automatic scaling, and finally saves the workbook as a PDF while preserving the original page dimensions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password provided)
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Ensure workbook was loaded successfully
            if (workbook == null)
            {
                Console.WriteLine("Error: Workbook could not be initialized.");
                return;
            }

            // Disable automatic scaling for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;
                pageSetup.Zoom = 100;               // 100% zoom (no scaling)
                pageSetup.FitToPagesWide = 0;       // Disable FitToPagesWide
                pageSetup.FitToPagesTall = 0;       // Disable FitToPagesTall
            }

            // Save the workbook as PDF preserving original dimensions
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing or saving: {ex.Message}");
        }
    }
}
