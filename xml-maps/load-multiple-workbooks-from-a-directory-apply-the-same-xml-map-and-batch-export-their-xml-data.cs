using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchExport
{
    public class XmlBatchExporter
    {
        /// <summary>
        /// Loads all Excel workbooks from the specified directory, uses the first XML map found in each workbook,
        /// and exports the mapped data to individual XML files in the output directory.
        /// </summary>
        /// <param name="inputFolder">Folder containing the source Excel files.</param>
        /// <param name="outputFolder">Folder where the exported XML files will be saved.</param>
        public static void Run(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all Excel files in the input folder (supports .xlsx, .xls, .xlsm)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Load the workbook using the constructor rule
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Verify that the workbook contains at least one XML map
                    if (workbook.Worksheets.XmlMaps.Count == 0)
                    {
                        Console.WriteLine($"No XML map found in '{Path.GetFileName(filePath)}'. Skipping.");
                        continue;
                    }

                    // Use the first XML map (you can adjust the selection logic as needed)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
                    string mapName = xmlMap.Name;

                    // Build the output XML file path
                    string xmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".xml";
                    string xmlOutputPath = Path.Combine(outputFolder, xmlFileName);

                    try
                    {
                        // Export XML using the ExportXml(string, string) rule
                        workbook.ExportXml(mapName, xmlOutputPath);
                        Console.WriteLine($"Exported XML for '{Path.GetFileName(filePath)}' to '{xmlFileName}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error exporting '{Path.GetFileName(filePath)}': {ex.Message}");
                    }
                }
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceFolder = @"C:\InputWorkbooks";
            string destinationFolder = @"C:\ExportedXml";

            XmlBatchExporter.Run(sourceFolder, destinationFolder);
        }
    }
}