// Title: Remove all slicers from every worksheet in an Excel file and save the workbook as a PDF using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook, iterate over each worksheet, delete every slicer from the worksheet's SlicerCollection, and then save the workbook as a PDF with only visible sheets using PdfSaveOptions. | Using Aspose.Cells for .NET, programmatically clear slicers on all sheets, handle a missing input file gracefully, and generate a PDF output of the cleaned workbook.
// Common Searches: C# Aspose.Cells delete slicers from all worksheets before PDF conversion | How to remove slicers in an Excel workbook using Aspose.Cells .NET | Export Excel to PDF without slicers using Aspose.Cells | Iterate worksheets and clear SlicerCollection in Aspose.Cells C# example | Aspose.Cells PDF save options to include only visible sheets after removing slicers
// Tags: remove slicers Aspose.Cells C# | export workbook to PDF Aspose.Cells | clear worksheet SlicerCollection .NET | PdfSaveOptions visible sheets Aspose.Cells | handle missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;   // Required for slicer classes

namespace DeleteSlicersAndExportPdf
{
    // The sample loads 'input.xlsx', checks for its existence, iterates through every worksheet to delete all slicers by clearing each sheet's SlicerCollection, and then saves the modified workbook as 'output.pdf' using PdfSaveOptions configured to export only visible sheets. Errors such as missing files are caught and reported.
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
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and remove any slicers
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the slicer collection for the current worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Remove all slicers – iterate backwards to avoid index shifting
                    for (int i = slicers.Count - 1; i >= 0; i--)
                    {
                        slicers.RemoveAt(i);
                    }
                }

                // Set PDF save options (export all visible sheets)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = SheetSet.Visible
                };

                // Save the modified workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
