// Title: C# – Remove All Slicers from an Excel Workbook and Export as a Single PDF with Aspose.Cells
// Description: Loads an Excel file, clears every slicer on each worksheet using the SlicerCollection, configures PdfSaveOptions to include all sheets, and saves the workbook as one consolidated PDF. The sample also checks for a missing source file and catches runtime exceptions.
// Keywords: Aspose.Cells remove slicers C# | Aspose.Cells export PDF | clear slicer collection | PdfSaveOptions all sheets | C# workbook to PDF | delete slicers before PDF conversion | Aspose.Cells SlicerCollection Clear | consolidated PDF from Excel | Aspose.Cells example GitHub | C# Excel slicer removal
// Common Searches: how to delete all slicers using Aspose.Cells C# | export Excel to single PDF after removing slicers | Aspose.Cells clear slicers code example | C# remove slicers and save workbook as PDF | batch remove slicers from multiple workbooks
// Developer Intent: Strip every slicer from a workbook and generate one combined PDF file.
// Use Cases: Create a printable version of a dashboard by removing interactive slicers before PDF conversion. | Automate archival of reports: delete slicers and store the clean workbook as a single PDF. | Prepare client‑ready documents where slicers are unnecessary, ensuring only data tables appear.
// AI Prompts: Generate C# code that uses Aspose.Cells to clear all slicers from every worksheet and save the workbook as a single PDF. | Explain how Aspose.Cells handles workbooks that contain no slicers when calling SlicerCollection.Clear(). | Suggest PdfSaveOptions settings to change page orientation, scaling, or add a header after slicers have been removed.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;   // Required for SlicerCollection

namespace RemoveSlicersAndSavePdf
{
    // Loads an Excel file, clears every slicer on each worksheet using the SlicerCollection, configures PdfSaveOptions to include all sheets, and saves the workbook as one consolidated PDF. The sample also checks for a missing source file and catches runtime exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Remove all slicers from each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    SlicerCollection slicers = sheet.Slicers;
                    slicers.Clear();
                }

                // Configure PDF save options (include all sheets)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = SheetSet.All
                };

                // Save the workbook as a consolidated PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
