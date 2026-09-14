// Title: Measure performance of batch saving Excel workbooks with Aspose.Cells using ExcludeUnusedStyles and ExportGridLines settings
// AI Prompts: Write a C# console program that iterates through all .xlsx files in a directory, loads each workbook with Aspose.Cells, saves it with the Workbook.Settings properties ExcludeUnusedStyles and ExportGridLines enabled via reflection, and records the total elapsed time. | Modify the batch processing loop to first save each workbook with default settings, then repeat the save using the custom settings, and output a side‑by‑side timing comparison. | Add robust error handling that logs the file path and exception message for any workbook that fails during the custom‑settings save while allowing the remaining files to continue processing.
// Common Searches: how to benchmark Aspose.Cells workbook save speed with custom settings | c# batch convert xlsx files using Aspose.Cells and enable ExcludeUnusedStyles | using reflection to set ExportGridLines property in Aspose.Cells Workbook.Settings | compare default save time versus ExcludeUnusedStyles and ExportGridLines in Aspose.Cells
// Tags: batch workbook save Aspose.Cells custom settings | exclude unused styles Aspose.Cells performance | export grid lines Aspose.Cells save option | reflection set Workbook.Settings C# | benchmark Aspose.Cells save time

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // // Scans a folder for .xlsx files, saves each workbook twice with Aspose.Cells—once using default options and once with ExcludeUnusedStyles and ExportGridLines enabled via reflection—while timing both batches and logging any errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\Workbooks\Input";
            // Folder to save processed workbooks
            string outputFolderDefault = @"C:\Workbooks\Output\Default";
            string outputFolderCustom = @"C:\Workbooks\Output\Custom";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directories exist
            Directory.CreateDirectory(outputFolderDefault);
            Directory.CreateDirectory(outputFolderCustom);

            // Gather all Excel files in the input folder
            List<string> workbookFiles = new List<string>(Directory.GetFiles(inputFolder, "*.xlsx"));

            // Measure processing time with default settings
            Stopwatch swDefault = Stopwatch.StartNew();
            foreach (string filePath in workbookFiles)
            {
                ProcessWorkbook(filePath, outputFolderDefault, applyCustomSettings: false);
            }
            swDefault.Stop();

            // Measure processing time with custom settings (if supported)
            Stopwatch swCustom = Stopwatch.StartNew();
            foreach (string filePath in workbookFiles)
            {
                ProcessWorkbook(filePath, outputFolderCustom, applyCustomSettings: true);
            }
            swCustom.Stop();

            // Output the timing results
            Console.WriteLine($"Processing {workbookFiles.Count} workbooks with default settings took: {swDefault.Elapsed.TotalSeconds:F2} seconds.");
            Console.WriteLine($"Processing {workbookFiles.Count} workbooks with custom settings took: {swCustom.Elapsed.TotalSeconds:F2} seconds.");
        }

        /// <param name="inputPath">Full path to the source workbook.</param>
        /// <param name="outputFolder">Folder where the processed workbook will be saved.</param>
        /// <param name="applyCustomSettings">If true, attempts to set custom workbook options.</param>
        private static void ProcessWorkbook(string inputPath, string outputFolder, bool applyCustomSettings)
        {
            // Verify the source file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                if (applyCustomSettings)
                {
                    // Use reflection to set properties that may not exist in the current Aspose.Cells version
                    SetWorkbookSettingIfExists(workbook.Settings, "ExcludeUnusedStyles", true);
                    SetWorkbookSettingIfExists(workbook.Settings, "ExportGridLines", true);
                }

                // Determine output file name and path
                string fileName = Path.GetFileName(inputPath);
                string outputPath = Path.Combine(outputFolder, fileName);

                // Save the workbook
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
            }
        }

        // Helper method to set a workbook setting via reflection if the property exists
        private static void SetWorkbookSettingIfExists(object settingsObject, string propertyName, object value)
        {
            try
            {
                PropertyInfo prop = settingsObject.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(settingsObject, value);
                }
            }
            catch (Exception ex)
            {
                // Log but ignore any reflection errors to keep processing robust
                Console.WriteLine($"Unable to set '{propertyName}': {ex.Message}");
            }
        }
    }
}
