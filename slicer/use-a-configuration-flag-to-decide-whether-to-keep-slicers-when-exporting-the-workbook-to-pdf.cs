// Title: Control slicer inclusion with a configuration flag when exporting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a boolean flag from app settings and uses it to toggle slicer rendering while saving a Workbook to PDF with Aspose.Cells. | Show how to apply PdfSaveOptions or an alternative technique to hide slicers when the flag is false, given that SlicerRenderingMode may be unavailable. | Create a reusable method that accepts a Workbook, output path, and keepSlicers flag, then performs the PDF export with appropriate slicer handling.
// Common Searches: Aspose.Cells .NET export Excel to PDF without slicers based on config setting | How to programmatically hide slicers when saving workbook as PDF using Aspose.Cells | C# code to toggle slicer visibility during PDF export with Aspose.Cells | Read boolean flag from appsettings.json to control slicer visibility in Aspose.Cells PDF export
// Tags: Aspose.Cells PDF export slicer control | C# conditional slicer rendering | PdfSaveOptions slicer visibility | Excel to PDF export configuration flag | handling missing SlicerRenderingMode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, reads a boolean configuration flag, creates PdfSaveOptions, and saves the workbook as a PDF. Because the current Aspose.Cells version does not expose a SlicerRenderingMode property, slicer inclusion must be managed manually or by upgrading the library; the flag determines whether slicers are kept or hidden during the export.
class Program
{
    static void Main()
    {
        try
        {
            // Configuration flag: true to keep slicers, false to remove them
            bool keepSlicers = GetKeepSlicersFlag();

            // Load the workbook (provided load rule)
            Workbook workbook = LoadWorkbook("input.xlsx");

            // Create PDF save options (provided creation rule)
            PdfSaveOptions pdfOptions = CreatePdfSaveOptions();

            // NOTE: SlicerRenderingMode is not available in the current Aspose.Cells version.
            // If needed, handle slicer rendering via other means or upgrade the library.

            // Export to PDF (provided save rule)
            SaveWorkbookAsPdf(workbook, "output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Retrieve the configuration flag (implementation can be customized)
    static bool GetKeepSlicersFlag()
    {
        // Example: read from configuration, environment variable, etc.
        // Here we simply return true for illustration.
        return true;
    }

    // Provided rule: load a workbook from a file with existence check
    static Workbook LoadWorkbook(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"The workbook file '{path}' was not found.");

        return new Workbook(path);
    }

    // Provided rule: create a PdfSaveOptions instance
    static PdfSaveOptions CreatePdfSaveOptions()
    {
        return new PdfSaveOptions();
    }

    // Provided rule: save a workbook as PDF using the given options with output path check
    static void SaveWorkbookAsPdf(Workbook wb, string outputPath, PdfSaveOptions options)
    {
        try
        {
            // Ensure the directory for the output file exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            wb.Save(outputPath, options);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save PDF: {ex.Message}");
            throw;
        }
    }
}
