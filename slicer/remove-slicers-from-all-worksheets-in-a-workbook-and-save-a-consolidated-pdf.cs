using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;   // Required for slicer classes

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that may contain slicers
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and remove every slicer
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the slicer collection for the current worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Clear all slicers if any exist
                    if (slicers != null && slicers.Count > 0)
                    {
                        slicers.Clear();
                    }
                }

                // Prepare PDF save options (default renders all visible sheets)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the cleaned workbook as a consolidated PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook processed successfully. PDF saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}