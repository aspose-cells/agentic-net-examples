// Title: Batch load Excel workbooks, apply a shared XML map, and export each workbook to XML using Aspose.Cells for .NET
// Common Searches: how to export multiple Excel files to XML with a single XmlMap using Aspose.Cells C# | C# code to load all .xlsx files from a folder and generate XML output with Aspose.Cells | apply the same XML schema to several workbooks before exporting to XML in Aspose.Cells | batch processing Excel workbooks to XML with Aspose.Cells .NET library | ensure an XmlMap exists in a workbook and export it to XML programmatically
// Tags: batch export xml from excel workbooks Aspose.Cells | shared xml map application across workbooks C# | load workbooks from directory Aspose.Cells | export workbook to xml file Aspose.Cells | add or reuse XmlMap before export Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchXmlExport
{
    // Loads every .xlsx/.xls/.xlsm file in a specified folder, adds the given XML schema as an XmlMap if it is not already present, and uses Aspose.Cells to export each workbook’s data to an XML file with the same base name in an output directory.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: BatchXmlExport <inputDir> <xmlSchemaPath> <outputDir>");
                return;
            }

            string inputDir = args[0];
            string xmlSchemaPath = args[1];
            string outputDir = args[2];

            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Error: Input directory does not exist: {inputDir}");
                return;
            }

            if (!File.Exists(xmlSchemaPath))
            {
                Console.WriteLine($"Error: XML schema file not found: {xmlSchemaPath}");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                try
                {
                    Directory.CreateDirectory(outputDir);
                    Console.WriteLine($"Created output directory: {outputDir}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating output directory: {ex.Message}");
                    return;
                }
            }

            // Gather all Excel files (supporting .xlsx, .xls, .xlsm)
            string[] excelFiles = Directory.GetFiles(inputDir, "*.*", SearchOption.TopDirectoryOnly);
            List<string> workbookFiles = new List<string>();
            foreach (var file in excelFiles)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm")
                {
                    workbookFiles.Add(file);
                }
            }

            if (workbookFiles.Count == 0)
            {
                Console.WriteLine("No Excel workbooks found in the specified input directory.");
                return;
            }

            Console.WriteLine($"Found {workbookFiles.Count} workbook(s). Starting batch processing...");

            foreach (var workbookPath in workbookFiles)
            {
                try
                {
                    Console.WriteLine($"\nProcessing workbook: {Path.GetFileName(workbookPath)}");

                    // Load the workbook
                    using (Workbook wb = new Workbook(workbookPath))
                    {
                        // Ensure the XML map exists (add if missing)
                        string mapName = EnsureXmlMap(wb, xmlSchemaPath);

                        // Determine output XML file path (same base name, .xml extension)
                        string outputXmlPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(workbookPath) + ".xml");

                        // Export XML using the map name
                        wb.ExportXml(mapName, outputXmlPath);
                        Console.WriteLine($"Exported XML to: {outputXmlPath}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{workbookPath}'. Error: {ex.Message}");
                }
            }

            Console.WriteLine("\nBatch export completed.");
        }

        /// <summary>
        /// Ensures that the workbook contains an XmlMap based on the provided schema.
        /// If a map with the same name already exists, it is reused; otherwise a new map is added.
        /// Returns the name of the XmlMap to be used for ExportXml.
        /// </summary>
        private static string EnsureXmlMap(Workbook workbook, string schemaPath)
        {
            // Load the schema content (Aspose expects a path or URL)
            // Adding the map returns its index; we will check for existing maps with the same name.
            // First, try to find an existing map that matches the schema file name (without extension).
            string candidateName = Path.GetFileNameWithoutExtension(schemaPath);

            foreach (XmlMap existingMap in workbook.Worksheets.XmlMaps)
            {
                if (string.Equals(existingMap.Name, candidateName, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Reusing existing XmlMap: {existingMap.Name}");
                    return existingMap.Name;
                }
            }

            // Map not found; add a new one.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap newMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Optionally set a friendly name if the default is not suitable.
            if (!string.Equals(newMap.Name, candidateName, StringComparison.OrdinalIgnoreCase))
            {
                newMap.Name = candidateName;
            }

            Console.WriteLine($"Added new XmlMap: {newMap.Name}");
            return newMap.Name;
        }
    }
}
