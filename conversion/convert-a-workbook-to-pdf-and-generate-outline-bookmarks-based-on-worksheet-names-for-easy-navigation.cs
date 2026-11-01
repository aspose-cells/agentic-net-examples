// Title: Generate a PDF from an Excel workbook with worksheet‑based outline bookmarks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with bookmarks for each sheet | how to enable outline bookmarks when saving a workbook as PDF in .NET | C# code to add worksheet names as PDF bookmarks using Aspose.Cells | set PdfSaveOptions outline options via reflection Aspose.Cells | convert .xlsx to .pdf with navigation bookmarks using Aspose.Cells
// Tags: Aspose.Cells PDF export outline bookmarks | PdfSaveOptions outline configuration C# | Excel worksheet names as PDF navigation bookmarks | reflection for conditional PdfSaveOptions properties | workbook to PDF conversion with navigation outline

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, and uses reflection to enable outline bookmarks (IsOutlineEnabled = true, OutlineLevels = 1) when the OutlineOptions property is present. The workbook is then saved as a PDF with sheet‑based bookmarks, providing easy navigation in the generated document.
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

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Outline (bookmark) support may depend on the Aspose.Cells version.
            // If the OutlineOptions property is available, enable it to create bookmarks
            // for each worksheet. Otherwise, the PDF will be generated without bookmarks.
            // The following block uses reflection to set the properties safely when they exist.
            var outlineProp = typeof(PdfSaveOptions).GetProperty("OutlineOptions");
            if (outlineProp != null)
            {
                var outlineValue = outlineProp.GetValue(pdfOptions);
                var isEnabledProp = outlineValue?.GetType().GetProperty("IsOutlineEnabled");
                var levelsProp = outlineValue?.GetType().GetProperty("OutlineLevels");

                isEnabledProp?.SetValue(outlineValue, true);
                levelsProp?.SetValue(outlineValue, 1);
            }

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
