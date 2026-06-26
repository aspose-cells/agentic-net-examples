using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchExport
{
    public class XmlBatchExporter
    {
        // Entry point for the batch export process
        public static void Run(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (non‑recursive)
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                // Guard against missing files (should not happen, but safe)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}. Skipping.");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(filePath);

                    // If the workbook contains no XML maps, skip it
                    if (workbook.Worksheets.XmlMaps.Count == 0)
                    {
                        Console.WriteLine($"No XmlMap found in '{Path.GetFileName(filePath)}'. Skipping.");
                        continue;
                    }

                    // Iterate through each XmlMap in the workbook
                    for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                    {
                        XmlMap xmlMap = workbook.Worksheets.XmlMaps[i];

                        // Build a unique output file name: <WorkbookName>_<MapName>.xml
                        string workbookName = Path.GetFileNameWithoutExtension(filePath);
                        string safeMapName = MakeFileSystemSafe(xmlMap.Name);
                        string outputFile = Path.Combine(outputFolder, $"{workbookName}_{safeMapName}.xml");

                        // Export the XML data using the map's name
                        workbook.ExportXml(xmlMap.Name, outputFile);
                        Console.WriteLine($"Exported XML map '{xmlMap.Name}' from '{workbookName}' to '{outputFile}'.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }

        // Helper to replace invalid filename characters
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }

    // Console entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: AsposeCellsBatchExport <inputFolder> <outputFolder>");
                    return;
                }

                string inputFolder = args[0];
                string outputFolder = args[1];

                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                XmlBatchExporter.Run(inputFolder, outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}