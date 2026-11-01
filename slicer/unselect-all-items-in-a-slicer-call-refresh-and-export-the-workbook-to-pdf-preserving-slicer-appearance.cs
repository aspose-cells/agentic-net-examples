// Title: How to deselect all slicer items, refresh slicers and pivot tables, and export an Excel workbook to PDF with slicer rendering using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates through every worksheet, clears the selection of each slicer item, calls Refresh on each slicer, refreshes all pivot tables, and saves the workbook as a PDF while preserving slicer appearance. | Provide an example that uses Aspose.Cells PdfSaveOptions to enable slicer rendering mode when exporting an Excel file to PDF after programmatically resetting slicer selections.
// Common Searches: Aspose.Cells C# clear all slicer selections before exporting to PDF | programmatically deselect slicer items in an Excel workbook using Aspose.Cells | refresh pivot tables after modifying slicer cache with Aspose.Cells .NET | export Excel to PDF with slicer rendering mode Aspose.Cells | preserve slicer appearance when saving workbook as PDF using Aspose.Cells
// Tags: Aspose.Cells clear slicer cache items | Aspose.Cells slicer refresh method | Aspose.Cells PDF export with slicer rendering | Aspose.Cells pivot table data refresh | Aspose.Cells workbook to PDF C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, loops through each worksheet to deselect every item in all slicers via the SlicerCache, calls Refresh on each slicer, refreshes all pivot tables, and then saves the workbook as a PDF using PdfSaveOptions (optionally setting SlicerRenderingMode to keep slicer visuals).
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Process slicers: clear selections and refresh.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Slicers != null && sheet.Slicers.Count > 0)
                {
                    foreach (Slicer slicer in sheet.Slicers)
                    {
                        try
                        {
                            // Clear selection of all slicer items if the API is available.
                            if (slicer.SlicerCache != null && slicer.SlicerCache.SlicerCacheItems != null)
                            {
                                foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                                {
                                    item.Selected = false;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log but continue if clearing fails.
                            Console.WriteLine($"Failed to clear slicer selections: {ex.Message}");
                        }

                        // Refresh the slicer to apply any changes.
                        slicer.Refresh();
                    }
                }
            }

            // Refresh all pivot tables (if any).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    pt.RefreshData();
                }
            }

            // Save the workbook as PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // If the Aspose.Cells version supports slicer rendering mode, uncomment the line below:
            // pdfOptions.SlicerRenderingMode = SlicerRenderingMode.Pdf;

            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
