using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsAsyncExport
{
    public class CertificateExporter
    {
        /// <summary>
        /// Exports the first XML map (treated as a certificate) from each workbook asynchronously.
        /// </summary>
        /// <param name="workbookPaths">Full paths of the workbooks to process.</param>
        /// <param name="outputDirectory">Directory where the exported XML files will be saved.</param>
        /// <returns>A task that represents the asynchronous operation.</returns>
        public static async Task ExportCertificatesAsync(string[] workbookPaths, string outputDirectory)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDirectory);

            var exportTasks = new List<Task>();

            foreach (string wbPath in workbookPaths)
            {
                // Capture the current path for the lambda.
                string currentPath = wbPath;

                // Create a task for each workbook.
                Task exportTask = Task.Run(() =>
                {
                    try
                    {
                        // Verify the workbook file exists before loading.
                        if (!File.Exists(currentPath))
                        {
                            Console.WriteLine($"File not found: {currentPath}");
                            return;
                        }

                        // Load the workbook.
                        using (var workbook = new Workbook(currentPath))
                        {
                            // Check if there is at least one XmlMap.
                            if (workbook.Worksheets.XmlMaps.Count > 0)
                            {
                                // Get the first XmlMap.
                                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                                // Build the output XML file name.
                                string xmlFileName = Path.GetFileNameWithoutExtension(currentPath) + "_" + xmlMap.Name + ".xml";
                                string xmlFullPath = Path.Combine(outputDirectory, xmlFileName);

                                // Export the XML data.
                                workbook.ExportXml(xmlMap.Name, xmlFullPath);
                            }
                            else
                            {
                                Console.WriteLine($"No XmlMap found in workbook: {currentPath}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log any unexpected errors for this workbook.
                        Console.WriteLine($"Error processing '{currentPath}': {ex.Message}");
                    }
                });

                exportTasks.Add(exportTask);
            }

            // Await all export tasks to complete.
            await Task.WhenAll(exportTasks);
        }

        // Example usage.
        public static async Task RunExample()
        {
            // Example list of workbook files.
            string[] workbooks = new string[]
            {
                @"C:\Data\Workbook1.xlsx",
                @"C:\Data\Workbook2.xlsx",
                @"C:\Data\Workbook3.xlsx"
            };

            // Destination folder for exported XML certificates.
            string outputDir = @"C:\Data\ExportedCertificates";

            // Perform asynchronous export.
            await ExportCertificatesAsync(workbooks, outputDir);

            Console.WriteLine("All certificates have been exported.");
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static async Task Main(string[] args)
        {
            await CertificateExporter.RunExample();
        }
    }
}