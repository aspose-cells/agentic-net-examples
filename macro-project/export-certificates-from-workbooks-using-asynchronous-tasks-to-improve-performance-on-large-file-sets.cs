// Title: Async Export of XML Map Certificates from Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: Shows how to load a collection of Excel files, detect embedded XML maps (treated as certificates), and export each map to its own XML file using parallel Tasks. The solution creates the output folder, runs workbook processing concurrently, logs missing files or errors, and awaits all tasks for high‑throughput batch handling.
// Keywords: Aspose.Cells | C# async export | XML map | certificate extraction | parallel workbook processing | batch Excel export | ExportXml | Task.WhenAll | Excel to XML | large file sets
// Common Searches: export xml maps from excel using aspose.cells c# | async batch export of excel certificates | parallel processing of multiple workbooks Aspose | how to extract xml map from workbook .NET | asynchronous export of excel xml schemas
// Developer Intent: Implement a scalable, asynchronous routine that extracts XML map certificates from many Excel workbooks in parallel to improve performance.
// Use Cases: Mass extraction of XML schemas from thousands of financial reports for downstream analytics. | Automated generation of per‑workbook XML files during a data‑migration pipeline. | Background service that monitors an upload folder and continuously pulls certificates from new workbooks. | Performance‑critical ETL job that needs to process large batches of Excel files without blocking the main thread.
// AI Prompts: Create a version of ExportCertificatesAsync that limits concurrency to a configurable maximum (e.g., four parallel tasks). | Add detailed logging with timestamps and write a summary report after all exports complete. | Write unit tests that mock workbooks with and without XML maps and verify the correct files are created and errors are handled. | Refactor the code to use Parallel.ForEach instead of manual Task collection while preserving async/await semantics.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsAsyncExport
{
    // Shows how to load a collection of Excel files, detect embedded XML maps (treated as certificates), and export each map to its own XML file using parallel Tasks. The solution creates the output folder, runs workbook processing concurrently, logs missing files or errors, and awaits all tasks for high‑throughput batch handling.
    public class CertificateExporter
    {
        // Export XML maps (treated as certificates) from a list of workbooks asynchronously.
        public static async Task ExportCertificatesAsync(string[] workbookPaths, string outputDirectory)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDirectory);

            // Create a list to hold the export tasks.
            List<Task> exportTasks = new List<Task>();

            foreach (string wbPath in workbookPaths)
            {
                // For each workbook start a separate task.
                exportTasks.Add(Task.Run(() =>
                {
                    // Verify the workbook file exists.
                    if (!File.Exists(wbPath))
                    {
                        Console.WriteLine($"File not found: '{wbPath}'. Skipping.");
                        return;
                    }

                    try
                    {
                        // Load the workbook.
                        using (Workbook workbook = new Workbook(wbPath))
                        {
                            // If there are no XML maps, nothing to export.
                            if (workbook.Worksheets.XmlMaps.Count == 0)
                            {
                                Console.WriteLine($"No XML maps found in '{wbPath}'.");
                                return;
                            }

                            // Export each XML map to a separate file.
                            for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                            {
                                var xmlMap = workbook.Worksheets.XmlMaps[i];
                                string mapName = xmlMap.Name;

                                // Build the output file name: <workbookName>_<mapName>.xml
                                string wbFileName = Path.GetFileNameWithoutExtension(wbPath);
                                string outputPath = Path.Combine(outputDirectory, $"{wbFileName}_{mapName}.xml");

                                // Export the XML map.
                                workbook.ExportXml(mapName, outputPath);
                                Console.WriteLine($"Exported map '{mapName}' from '{wbPath}' to '{outputPath}'.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing '{wbPath}': {ex.Message}");
                    }
                }));
            }

            // Await all export tasks to complete.
            await Task.WhenAll(exportTasks);
        }

        // Example usage.
        public static async Task RunDemo()
        {
            // Example list of workbook files (replace with actual paths).
            string[] workbooks = new[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            string outputDir = @"C:\Data\ExportedCertificates";

            await ExportCertificatesAsync(workbooks, outputDir);
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static async Task Main(string[] args)
        {
            await CertificateExporter.RunDemo();
        }
    }
}
