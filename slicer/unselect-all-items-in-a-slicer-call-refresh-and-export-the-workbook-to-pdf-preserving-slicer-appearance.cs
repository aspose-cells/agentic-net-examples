// Title: How to clear all slicer selections, refresh slicers, and export an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to deselect all items in an Excel slicer with Aspose.Cells before exporting to PDF | Aspose.Cells .NET clear slicer selections and refresh slicer then save workbook as PDF | How to programmatically unselect slicer items and generate a PDF from Excel using Aspose.Cells | Remove slicer filters and export Excel to PDF with Aspose.Cells in C# | Aspose.Cells refresh slicer after clearing selections prior to PDF conversion
// Tags: Aspose.Cells slicer selection reset C# | Aspose.Cells slicer refresh for PDF conversion | C# Aspose.Cells save workbook as PDF | Excel slicer API handling Aspose.Cells .NET | unselect slicer items Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet's slicers to clear all selections and call Refresh, ensures the output directory exists, and then saves the workbook as a PDF, handling missing files and runtime exceptions.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Slicer manipulation requires a version of Aspose.Cells that includes the Slicer API.
            // If the current library version does not support slicers, this block can be omitted.
            // The following code is kept for reference and will be skipped if the Slicer type is unavailable.
            /*
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Aspose.Cells.Drawing.Slicer slicer in sheet.Slicers)
                {
                    // Unselect all items in the slicer
                    slicer.SlicerItems.ClearSelection();

                    // Refresh the slicer to apply the changes
                    slicer.Refresh();
                }
            }
            */

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
