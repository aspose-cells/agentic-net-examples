using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // SlicerCollection resides in this namespace
using Aspose.Cells.Rendering;

namespace AsposeCellsParallelSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook files (adjust paths as needed)
                string[] inputFiles = new string[]
                {
                    @"C:\Input\Workbook1.xlsx",
                    @"C:\Input\Workbook2.xlsx",
                    @"C:\Input\Workbook3.xlsx"
                };

                // Output PDF files (same name, different folder)
                string outputFolder = @"C:\Output\";
                Directory.CreateDirectory(outputFolder);

                // Prepare tasks for parallel processing
                List<Task> tasks = new List<Task>();
                foreach (string inputPath in inputFiles)
                {
                    // Verify source file exists before scheduling the task
                    if (!File.Exists(inputPath))
                    {
                        Console.WriteLine($"Source file not found: {inputPath}");
                        continue;
                    }

                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                    string outputPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    tasks.Add(Task.Run(() => ProcessWorkbook(inputPath, outputPath)));
                }

                // Wait for all tasks to complete
                Task.WaitAll(tasks.ToArray());

                Console.WriteLine("All workbooks processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads a workbook, removes the first slicer from each worksheet (if any),
        /// and saves the result as a PDF.
        /// </summary>
        /// <param name="inputPath">Path to the source Excel file.</param>
        /// <param name="outputPath">Path where the PDF will be saved.</param>
        static void ProcessWorkbook(string inputPath, string outputPath)
        {
            try
            {
                // Load the workbook from file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the slicer collection of the worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // If there is at least one slicer, remove the first one
                    if (slicers != null && slicers.Count > 0)
                    {
                        slicers.RemoveAt(0);
                    }
                }

                // Prepare PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the modified workbook as PDF
                workbook.Save(outputPath, pdfOptions);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
            }
        }
    }
}