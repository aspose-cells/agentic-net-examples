// Title: Asynchronous conversion of an Excel workbook with WordArt shapes to PDF using Aspose.Cells and progress reporting
// Common Searches: c# aspnet async convert excel workbook containing WordArt to pdf with Aspose.Cells | how to report conversion progress while exporting Excel to PDF using IProgress in Aspose.Cells | async pdf export of Excel file with WordArt shapes using Aspose.Cells PdfSaveOptions | simulate gradient rendering for WordArt shapes during async Excel to PDF conversion | use Task.Run to save Aspose.Cells workbook as PDF on a background thread
// Tags: async workbook to pdf conversion aspose.cells | wordart shape handling aspose.cells | pdfsaveoptions gradient rendering aspose.cells | iprogress progress reporting aspose.cells | task.run background pdf save aspose.cells

using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WordArtPdfAsyncConversion
{
    // The example loads an Excel workbook, counts all shapes, iterates through each worksheet and shape, simulates processing of WordArt shapes with a brief delay, reports conversion progress via an IProgress<int> instance, and finally saves the workbook as a PDF using PdfSaveOptions on a background thread with Task.Run, updating progress after the save operation.
    class Program
    {
        // Entry point - async Main is supported in C# 7.1+
        static async Task Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath = args.Length > 0 ? args[0] : "WordArtSample.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "WordArtSample.pdf";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: {inputPath}");

                // Prepare a progress reporter
                var progress = new Progress<int>(percent =>
                {
                    Console.WriteLine($"Conversion progress: {percent}%");
                });

                // Start asynchronous conversion
                await ConvertWorkbookToPdfAsync(workbook, outputPath, progress);
                Console.WriteLine($"PDF saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="workbook">The loaded Aspose.Cells Workbook.</param>
        /// <param name="outputPdfPath">Destination PDF file path.</param>
        /// <param name="progress">IProgress instance for reporting integer percentages.</param>
        /// <returns>Task representing the asynchronous operation.</returns>
        private static async Task ConvertWorkbookToPdfAsync(Workbook workbook, string outputPdfPath, IProgress<int> progress)
        {
            // Count total steps: each shape (for gradient simulation) + final save step
            int totalShapes = CountWordArtShapes(workbook);
            int totalSteps = totalShapes + 1; // +1 for the final PDF save
            int completedSteps = 0;

            // Simulate processing of each WordArt shape (e.g., gradient rendering)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsWordArt)
                    {
                        // Here you could manipulate gradient fills or other rendering options.
                        // For demonstration, we just pause briefly to simulate work.
                        await Task.Delay(100); // Simulated processing time

                        // Optionally, you could log shape details
                        Console.WriteLine($"Processed WordArt shape: \"{shape.Name}\" on sheet \"{sheet.Name}\"");
                    }

                    // Increment progress after each shape (regardless of WordArt status)
                    completedSteps++;
                    int percent = (int)((completedSteps / (double)totalSteps) * 100);
                    progress.Report(percent);
                }
            }

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example options – adjust as needed
                OnePagePerSheet = true,
                Compliance = PdfCompliance.PdfA1b,
                // You can enable gradient rendering options here if required
                // For example, set RenderingOptions if you need fine‑grained control
            };

            // Perform the actual save on a background thread
            await Task.Run(() =>
            {
                // Save the workbook as PDF
                workbook.Save(outputPdfPath, pdfOptions);
            });

            // Final progress update
            completedSteps++;
            int finalPercent = (int)((completedSteps / (double)totalSteps) * 100);
            progress.Report(finalPercent);
        }

        /// <summary>
        /// Counts the total number of shapes (including non‑WordArt) in the workbook.
        /// This count is used to calculate progress steps.
        /// </summary>
        /// <param name="workbook">The workbook to inspect.</param>
        /// <returns>Number of shapes across all worksheets.</returns>
        private static int CountWordArtShapes(Workbook workbook)
        {
            int count = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                count += sheet.Shapes.Count;
            }
            return count;
        }
    }
}
