// Title: Batch convert Excel workbooks to HTML with ExcludeUnusedStyles, ExportGridLines and benchmark performance
// Description: Scans an "InputWorkbooks" folder, loads each .xlsx file with Aspose.Cells, saves a default HTML version, then saves a custom HTML version using HtmlSaveOptions (ExcludeUnusedStyles = true, ExportGridLines = true), measures the elapsed time for each save, and outputs a side‑by‑side performance comparison.
// Keywords: Aspose.Cells batch HTML conversion | HtmlSaveOptions ExcludeUnusedStyles | ExportGridLines performance | Excel to HTML benchmark | measure Aspose.Cells save time | bulk workbook processing C#
// Common Searches: how to convert multiple Excel files to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles example | compare default and custom HTML save speed Aspose | batch export Excel to HTML with grid lines | measure Aspose.Cells HTML save time per workbook
// Developer Intent: Convert a collection of Excel workbooks to HTML with specific styling options and evaluate the impact on save speed versus the default configuration.
// Use Cases: Generate lightweight HTML reports for a large set of spreadsheets by omitting unused CSS. | Produce web‑ready HTML that preserves Excel grid lines for clearer visual layout. | Run performance benchmarks to decide whether custom HtmlSaveOptions affect processing time in bulk conversions.
// AI Prompts: Refactor the code to write timing results to a CSV file with columns: workbook, default_ms, custom_ms. | Show how to parallelize the conversion loop using Task Parallel Library while keeping accurate per‑file timing. | Explain how to disable ExportGridLines in HtmlSaveOptions and compare the resulting HTML file sizes.

using System;
using System.IO;
using System.Diagnostics;
using Aspose.Cells;

// Scans an "InputWorkbooks" folder, loads each .xlsx file with Aspose.Cells, saves a default HTML version, then saves a custom HTML version using HtmlSaveOptions (ExcludeUnusedStyles = true, ExportGridLines = true), measures the elapsed time for each save, and outputs a side‑by‑side performance comparison.
class BatchProcessWorkbooks
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = "InputWorkbooks";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder \"{inputFolder}\" not found. Please ensure the folder exists and contains .xlsx files.");
            return;
        }

        // Output folders for default and custom HTML saves
        string outputFolderDefault = "OutputDefault";
        string outputFolderCustom = "OutputCustom";

        // Ensure output directories exist
        Directory.CreateDirectory(outputFolderDefault);
        Directory.CreateDirectory(outputFolderCustom);

        // Get all .xlsx files in the input folder
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string workbookPath in workbookFiles)
        {
            // Skip if the file somehow does not exist
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook (create + load lifecycle)
                Workbook workbook = new Workbook(workbookPath);

                string fileBaseName = Path.GetFileNameWithoutExtension(workbookPath);

                // ---------- Default save (no special options) ----------
                string defaultHtmlPath = Path.Combine(outputFolderDefault, fileBaseName + "_default.html");
                Stopwatch swDefault = Stopwatch.StartNew();
                workbook.Save(defaultHtmlPath, SaveFormat.Html); // save lifecycle
                swDefault.Stop();

                // ---------- Custom save with ExcludeUnusedStyles & ExportGridLines ----------
                string customHtmlPath = Path.Combine(outputFolderCustom, fileBaseName + "_custom.html");
                HtmlSaveOptions customOptions = new HtmlSaveOptions
                {
                    ExcludeUnusedStyles = true,   // explicitly set (default is true)
                    ExportGridLines = true        // enable grid line export
                };
                Stopwatch swCustom = Stopwatch.StartNew();
                workbook.Save(customHtmlPath, customOptions); // save lifecycle with options
                swCustom.Stop();

                // Output timing comparison
                Console.WriteLine($"{fileBaseName}: Default = {swDefault.ElapsedMilliseconds} ms, Custom = {swCustom.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                // Log the error and continue processing other files
                Console.WriteLine($"Error processing \"{workbookPath}\": {ex.Message}");
            }
        }
    }
}
