using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions
using Aspose.Cells.Slicers;   // For SlicerCollection

namespace AsposeCellsParallelSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // List of input Excel files to process
            List<string> inputFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Corresponding output PDF files
            List<string> outputFiles = new List<string>
            {
                "Workbook1.pdf",
                "Workbook2.pdf",
                "Workbook3.pdf"
            };

            // Index of the slicer to remove from each worksheet (0‑based)
            int slicerIndexToRemove = 0;

            // Process each workbook in parallel
            Parallel.ForEach(
                Enumerable.Range(0, inputFiles.Count),
                index =>
                {
                    try
                    {
                        string inputPath = inputFiles[index];
                        string outputPath = outputFiles[index];
                        ProcessWorkbook(inputPath, outputPath, slicerIndexToRemove);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file pair #{index}: {ex.Message}");
                    }
                });
        }

        /// <summary>
        /// Loads a workbook, removes the slicer at the specified index from every worksheet,
        /// and saves the result as a PDF.
        /// </summary>
        /// <param name="inputPath">Path to the source Excel file.</param>
        /// <param name="outputPath">Path where the PDF will be saved.</param>
        /// <param name="slicerIndex">Zero‑based index of the slicer to remove.</param>
        private static void ProcessWorkbook(string inputPath, string outputPath, int slicerIndex)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the slicer collection of the worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Remove the slicer if the collection contains enough items
                    if (slicers.Count > slicerIndex)
                    {
                        slicers.RemoveAt(slicerIndex);
                    }
                }

                // Prepare PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Successfully saved PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to process workbook '{inputPath}': {ex.Message}");
            }
        }
    }
}