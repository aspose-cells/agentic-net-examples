// Title: Batch export XML from multiple Excel workbooks using a shared XML map – Aspose.Cells for .NET
// Description: Loads every .xlsx file in a given folder, ensures each workbook contains the XML map defined by a supplied XSD, adds the map when missing, and exports the workbook data to individual XML files with Aspose.Cells' ExportXml method.
// Keywords: Aspose.Cells batch XML export | C# Excel to XML | apply XML map to multiple workbooks | export workbook as XML | XSD schema Excel mapping | process folder of xlsx files | ExportXml Aspose.Cells | automated XML generation from Excel
// Common Searches: How to export XML from all Excel files in a directory using Aspose.Cells | Add missing XML map to workbooks before exporting with Aspose.Cells | Batch convert .xlsx to .xml with a shared XSD schema | Aspose.Cells export multiple workbooks to XML programmatically | C# code to apply an XML map to many Excel files
// Developer Intent: Create XML files for every Excel workbook in a folder by applying a common XSD map and exporting with Aspose.Cells.
// Use Cases: Nightly automation that converts a set of Excel templates into XML reports required by an external system. | Migration of legacy spreadsheet data to an XML format while guaranteeing each file contains the correct XML map. | Bulk validation and correction of workbooks lacking the required map before downstream XML processing.
// AI Prompts: Generate C# code that scans a directory for .xlsx files, adds a missing XML map from a specified XSD using Aspose.Cells, and exports each workbook to an XML file. | Explain best practices for error handling when loading workbooks or exporting XML in a batch operation with Aspose.Cells. | Show how to log the success or failure of each file during a bulk XML export using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchXmlExport
{
    // Loads every .xlsx file in a given folder, ensures each workbook contains the XML map defined by a supplied XSD, adds the map when missing, and exports the workbook data to individual XML files with Aspose.Cells' ExportXml method.
    public class BatchExporter
    {
        /// <param name="inputDirectory">Folder containing the source Excel files.</param>
        /// <param name="xmlSchemaPath">Path to the XSD/XML schema file that defines the XML map.</param>
        /// <param name="outputDirectory">Folder where the exported XML files will be saved.</param>
        public static void Run(string inputDirectory, string xmlSchemaPath, string outputDirectory)
        {
            // Validate input parameters
            if (string.IsNullOrWhiteSpace(inputDirectory) || !Directory.Exists(inputDirectory))
                throw new DirectoryNotFoundException($"Input directory not found: {inputDirectory}");

            if (string.IsNullOrWhiteSpace(xmlSchemaPath) || !File.Exists(xmlSchemaPath))
                throw new FileNotFoundException($"XML schema file not found: {xmlSchemaPath}");

            // Ensure the output folder exists
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Get all Excel files in the input folder
            string[] excelFiles = Directory.GetFiles(inputDirectory, "*.xlsx");

            foreach (string excelPath in excelFiles)
            {
                // Extra safety check (files come from GetFiles, but guard against race conditions)
                if (!File.Exists(excelPath))
                    continue;

                try
                {
                    // Load the workbook
                    using (Workbook workbook = new Workbook(excelPath))
                    {
                        // Find existing XML map that matches the schema name (without extension)
                        XmlMap targetMap = null;
                        string expectedMapName = Path.GetFileNameWithoutExtension(xmlSchemaPath);
                        foreach (XmlMap map in workbook.Worksheets.XmlMaps)
                        {
                            if (string.Equals(map.Name, expectedMapName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetMap = map;
                                break;
                            }
                        }

                        // Add the map if it does not exist
                        if (targetMap == null)
                        {
                            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaPath);
                            targetMap = workbook.Worksheets.XmlMaps[mapIndex];
                        }

                        // Build the output XML file path
                        string xmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".xml";
                        string xmlOutputPath = Path.Combine(outputDirectory, xmlFileName);

                        // Export the XML data
                        workbook.ExportXml(targetMap.Name, xmlOutputPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point of the application.
        /// Expected arguments: <inputDirectory> <xmlSchemaPath> <outputDirectory>
        /// </summary>
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length != 3)
                {
                    Console.WriteLine("Usage: AsposeCellsBatchXmlExport <inputDirectory> <xmlSchemaPath> <outputDirectory>");
                    return;
                }

                string inputDir = args[0];
                string schemaPath = args[1];
                string outputDir = args[2];

                BatchExporter.Run(inputDir, schemaPath, outputDir);
                Console.WriteLine("Batch export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}
