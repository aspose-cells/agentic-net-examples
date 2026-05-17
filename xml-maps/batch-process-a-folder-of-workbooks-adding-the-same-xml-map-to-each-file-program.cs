using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace BatchXmlMapProcessor
{
    public static class WorkbookXmlMapBatcher
    {
        /// <summary>
        /// Adds the same XML map to every workbook in the specified folder.
        /// </summary>
        /// <param name="sourceFolder">Folder containing the source Excel files.</param>
        /// <param name="xmlMapSource">
        /// Path to an XSD/XML schema file or a raw XML schema string that will be added as the map.
        /// </param>
        /// <param name="outputFolder">
        /// Folder where the processed workbooks will be saved. If it does not exist, it will be created.
        /// </param>
        public static void ProcessFolder(string sourceFolder, string xmlMapSource, string outputFolder)
        {
            // Validate input folders
            if (!Directory.Exists(sourceFolder))
                throw new DirectoryNotFoundException($"Source folder not found: {sourceFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Define the file extensions we want to process
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Enumerate all matching files in the source folder (non‑recursive)
            var workbookFiles = Directory.GetFiles(sourceFolder)
                                         .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));

            foreach (var filePath in workbookFiles)
            {
                // Load the workbook using the string‑path constructor (rule: Workbook(string))
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Add the XML map to the workbook (rule: XmlMapCollection.Add(string))
                    // The Add method returns the index of the newly added map.
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlMapSource);

                    // Optional: give the map a friendly name (useful for later reference)
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                    xmlMap.Name = "SharedXmlMap";

                    // Determine the output file path (overwrite the original file or write to a new location)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the modified workbook (rule: Workbook.Save(string))
                    workbook.Save(outputPath);
                }
            }
        }

        // Example usage
        public static void Main()
        {
            // Folder containing the Excel files to process
            string sourceFolder = @"C:\InputWorkbooks";

            // Path to the XSD schema that defines the XML map (same for all workbooks)
            string xmlSchemaPath = @"C:\Schemas\MyMapSchema.xsd";

            // Folder where the updated workbooks will be written
            string outputFolder = @"C:\ProcessedWorkbooks";

            ProcessFolder(sourceFolder, xmlSchemaPath, outputFolder);

            Console.WriteLine("Batch processing completed.");
        }
    }
}