// Title: Parallel Worksheet‑to‑PNG Conversion with Aspose.Cells and TPL (C#)
// Description: Loads an Excel workbook, creates an output folder, and uses the Task Parallel Library to render each worksheet to PNG files concurrently. Each task configures ImageOrPrintOptions (PNG, OnePagePerSheet), employs SheetRender to export every page, disposes resources, and finally waits for all tasks before optionally saving the original workbook.
// Keywords: Aspose.Cells parallel rendering | C# Excel to PNG multi‑threaded | Task Parallel Library image export | SheetRender concurrent conversion | OnePagePerSheet PNG | multi‑core Excel image generation | .NET workbook to PNG | high‑performance Excel preview
// Common Searches: how to convert multiple Excel sheets to PNG in parallel C# | Aspose.Cells TPL example for sheet rendering | parallel image export from workbook using Aspose.Cells | improve Excel to PNG conversion speed .NET | concurrent SheetRender usage Aspose
// Developer Intent: Export every worksheet of an Excel file to separate PNG images simultaneously, leveraging all available CPU cores for faster throughput.
// Use Cases: Generate preview images for each sheet of large workbooks in a web service without blocking the request thread. | Batch‑process thousands of spreadsheets in a document‑management pipeline, creating per‑sheet thumbnails in minutes. | Run a background job on Azure VMs or on‑prem servers that reduces total conversion time by using all logical processors.
// AI Prompts: Rewrite the sample to limit parallelism to the machine's logical processor count using ParallelOptions. | Add robust error handling: capture exceptions inside each task, log them, and report a summary after Task.WaitAll. | Show a version that uses Parallel.ForEach instead of manually building a List<Task> for worksheet rendering.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsParallelRender
{
    // Loads an Excel workbook, creates an output folder, and uses the Task Parallel Library to render each worksheet to PNG files concurrently. Each task configures ImageOrPrintOptions (PNG, OnePagePerSheet), employs SheetRender to export every page, disposes resources, and finally waits for all tasks before optionally saving the original workbook.
    class Program
    {
        static void Main()
        {
            // Input Excel file
            string inputFile = "input.xlsx";

            // Output directory for PNG images
            string outputDir = "output_png";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputFile);

            // Prepare a list to hold rendering tasks
            List<Task> renderTasks = new List<Task>();

            // Iterate over each worksheet in the workbook
            for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
            {
                // Capture the current index for the task closure
                int currentSheetIdx = sheetIdx;

                // Create a task that renders the current worksheet
                Task task = Task.Run(() =>
                {
                    // Access the worksheet
                    Worksheet sheet = workbook.Worksheets[currentSheetIdx];

                    // Configure image rendering options (PNG, one page per sheet)
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,
                        OnePagePerSheet = true
                    };

                    // Initialize SheetRender (constructor rule)
                    SheetRender sheetRender = new SheetRender(sheet, options);

                    // Render each page of the worksheet to a separate PNG file
                    for (int pageIdx = 0; pageIdx < sheetRender.PageCount; pageIdx++)
                    {
                        string fileName = Path.Combine(
                            outputDir,
                            $"Sheet{currentSheetIdx}_Page{pageIdx}.png");

                        // Render page to file (ToImage overload rule)
                        sheetRender.ToImage(pageIdx, fileName);
                    }

                    // Release resources used by SheetRender
                    sheetRender.Dispose();
                });

                renderTasks.Add(task);
            }

            // Wait for all rendering tasks to complete
            Task.WaitAll(renderTasks.ToArray());

            // Optionally, save the original workbook for reference (save rule)
            string savedWorkbookPath = Path.Combine(outputDir, "original_workbook.xlsx");
            workbook.Save(savedWorkbookPath);

            Console.WriteLine("All worksheets have been rendered to PNG files.");
        }
    }
}
