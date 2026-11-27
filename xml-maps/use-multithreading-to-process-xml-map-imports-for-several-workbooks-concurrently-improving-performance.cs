// Title: Concurrent XML map import into multiple Excel workbooks using Aspose.Cells for .NET
// Common Searches: how to import XML data into several Excel workbooks in parallel with Aspose.Cells C# | parallel processing of .xlsx files using Aspose.Cells ImportXml and Task.Run | multithreaded batch import of XML maps into workbooks with Aspose.Cells .NET | C# example for concurrent ImportXml on a folder of Excel files | using Aspose.Cells to import XML into multiple workbooks concurrently
// Tags: aspocells importxml parallel processing | multithreaded workbook xml map import | c# task.run aspocells batch import | xlsx xml map concurrent import | parallel excel file processing aspocells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace XmlMapConcurrentImport
{
    // The sample loads all supported Excel files from a folder, matches each workbook with an XML file of the same name, and launches a separate Task for each pair. Each task uses Aspose.Cells to open the workbook, ensures a target worksheet, imports the XML data via ImportXml starting at cell A1, recalculates formulas, and saves the result as an .xlsx file in the output folder, while handling any exceptions from the parallel execution.
    class Program
    {
        static void Main(string[] args)
        {
            // Resolve folders
            string workbookFolder = args.Length > 0 ? args[0] : Path.Combine(Directory.GetCurrentDirectory(), "InputWorkbooks");
            string xmlFolder = args.Length > 1 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "InputXml");
            string outputFolder = args.Length > 2 ? args[2] : Path.Combine(Directory.GetCurrentDirectory(), "OutputWorkbooks");

            Console.WriteLine("Workbook folder : " + workbookFolder);
            Console.WriteLine("XML folder      : " + xmlFolder);
            Console.WriteLine("Output folder   : " + outputFolder);
            Console.WriteLine();

            // Validate folders
            if (!Directory.Exists(workbookFolder))
            {
                Console.WriteLine($"Error: Workbook folder does not exist: {workbookFolder}");
                return;
            }

            if (!Directory.Exists(xmlFolder))
            {
                Console.WriteLine($"Error: XML folder does not exist: {xmlFolder}");
                return;
            }

            Directory.CreateDirectory(outputFolder);

            // Gather workbook files (support .xlsx and .xls)
            string[] workbookFiles = Directory.GetFiles(workbookFolder, "*.*", SearchOption.TopDirectoryOnly);
            var supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".xlsx", ".xls", ".xlsm", ".xlsb" };
            var tasks = new List<Task>();

            foreach (string workbookPath in workbookFiles)
            {
                if (!supportedExtensions.Contains(Path.GetExtension(workbookPath)))
                {
                    Console.WriteLine($"Skipping unsupported file: {Path.GetFileName(workbookPath)}");
                    continue;
                }

                // Determine matching XML file (same base name, .xml extension)
                string baseName = Path.GetFileNameWithoutExtension(workbookPath);
                string xmlPath = Path.Combine(xmlFolder, baseName + ".xml");

                if (!File.Exists(xmlPath))
                {
                    Console.WriteLine($"Warning: No matching XML file for workbook '{baseName}'. Skipping.");
                    continue;
                }

                // Prepare output path
                string outputPath = Path.Combine(outputFolder, baseName + "_Processed.xlsx");

                // Launch processing in a separate task
                tasks.Add(Task.Run(() => ProcessWorkbook(workbookPath, xmlPath, outputPath)));
            }

            // Wait for all tasks to finish
            try
            {
                Task.WaitAll(tasks.ToArray());
                Console.WriteLine();
                Console.WriteLine("All workbook processing completed.");
            }
            catch (AggregateException aggEx)
            {
                Console.WriteLine("One or more errors occurred during processing:");
                foreach (var ex in aggEx.InnerExceptions)
                {
                    Console.WriteLine(" - " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Imports an XML file into a workbook using Aspose.Cells and saves the result.
        /// </summary>
        /// <param name="workbookPath">Full path to the source workbook.</param>
        /// <param name="xmlPath">Full path to the XML file to import.</param>
        /// <param name="outputPath">Full path where the processed workbook will be saved.</param>
        private static void ProcessWorkbook(string workbookPath, string xmlPath, string outputPath)
        {
            try
            {
                Console.WriteLine($"[Thread {Task.CurrentId}] Processing workbook: {Path.GetFileName(workbookPath)}");

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Ensure the target sheet exists; create if necessary
                string targetSheetName = "Sheet1";
                Worksheet worksheet = workbook.Worksheets[targetSheetName] ?? workbook.Worksheets[0];
                worksheet.Name = targetSheetName; // rename first sheet if it was different

                // Import XML data starting at cell A1 (row 0, column 0)
                // This method will also create an XmlMap internally if needed.
                workbook.ImportXml(xmlPath, targetSheetName, 0, 0);

                // Optional: recalculate formulas after import
                workbook.CalculateFormula();

                // Save the processed workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"[Thread {Task.CurrentId}] Saved processed workbook to: {Path.GetFileName(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Thread {Task.CurrentId}] Error processing '{Path.GetFileName(workbookPath)}': {ex.Message}");
                // Rethrow to propagate to the AggregateException in Main
                throw;
            }
        }
    }
}
