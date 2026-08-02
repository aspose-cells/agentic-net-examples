using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsAsyncExport
{
    class Program
    {
        // Entry point
        static async Task Main(string[] args)
        {
            // Example input files (replace with actual paths)
            var excelFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Output directory for exported XML files
            string outputDir = "ExportedXml";
            Directory.CreateDirectory(outputDir);

            // Run the export operation asynchronously for all workbooks
            await ExportXmlFromWorkbooksAsync(excelFiles, outputDir);

            Console.WriteLine("All export tasks completed.");
        }

        /// <summary>
        /// Exports the first XML map of each workbook to a separate XML file using asynchronous tasks.
        /// </summary>
        /// <param name="workbookPaths">List of Excel file paths.</param>
        /// <param name="outputDirectory">Directory where XML files will be saved.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private static async Task ExportXmlFromWorkbooksAsync(IEnumerable<string> workbookPaths, string outputDirectory)
        {
            // Create a list to hold the export tasks
            var exportTasks = new List<Task>();

            foreach (var path in workbookPaths)
            {
                // For each workbook, start a new task
                exportTasks.Add(Task.Run(() =>
                {
                    // Ensure the workbook file exists
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"File not found: {path}");
                        return;
                    }

                    // Load the workbook (uses the Workbook(string) constructor)
                    using (Workbook workbook = new Workbook(path))
                    {
                        // Check if any XML maps are defined
                        if (workbook.Worksheets.XmlMaps.Count > 0)
                        {
                            // Use the first XmlMap
                            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                            // Build the output XML file name
                            string xmlFileName = Path.GetFileNameWithoutExtension(path) + "_" + xmlMap.Name + ".xml";
                            string xmlPath = Path.Combine(outputDirectory, xmlFileName);

                            // Export the XML data (uses ExportXml(string, string))
                            workbook.ExportXml(xmlMap.Name, xmlPath);

                            Console.WriteLine($"Exported XML for '{path}' to '{xmlPath}'.");
                        }
                        else
                        {
                            Console.WriteLine($"No XmlMap found in workbook: {path}");
                        }
                    }
                }));
            }

            // Await completion of all export tasks
            await Task.WhenAll(exportTasks);
        }
    }
}