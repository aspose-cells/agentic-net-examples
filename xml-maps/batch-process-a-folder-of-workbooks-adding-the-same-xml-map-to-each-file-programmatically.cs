// Title: How to batch add the same XML map to multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# script to add an XSD as XML map to every .xlsx file in a folder with Aspose.Cells | automate adding XML map to a batch of Excel workbooks using Aspose.Cells .NET | Aspose.Cells bulk insert XML map from XSD into existing workbooks | command‑line tool for adding XML map to multiple Excel files in C#
// Tags: aspocells bulk xmlmap insertion | xmlmapcollection add xsd batch | c# process excel folder xml schema | overwrite excel workbooks with xml map | command line xml map batch processing

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapBatchProcessor
{
    // The console application validates a workbook folder path and an XSD schema file, creates a "Processed" subdirectory, and iterates over all Excel files (.xlsx, .xls, .xlsm, .xlsb) in the folder. For each workbook it loads the file with Aspose.Cells, checks the Worksheets.XmlMaps collection for an existing map with the specified name, adds the XML map from the XSD when missing, assigns a friendly name, and saves the modified workbook to the output folder while logging successes and failures.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Insufficient arguments.");
                Console.WriteLine("Usage: XmlMapBatchProcessor.exe <WorkbookFolderPath> <XmlSchemaFilePath> [MapName]");
                return;
            }

            string workbookFolder = args[0];
            string xmlSchemaPath = args[1];
            string desiredMapName = args.Length >= 3 ? args[2] : Path.GetFileNameWithoutExtension(xmlSchemaPath);

            // Validate folder
            if (!Directory.Exists(workbookFolder))
            {
                Console.WriteLine($"Error: The folder \"{workbookFolder}\" does not exist.");
                return;
            }

            // Validate XML schema file
            if (!File.Exists(xmlSchemaPath))
            {
                Console.WriteLine($"Error: The XML schema file \"{xmlSchemaPath}\" does not exist.");
                return;
            }

            // Prepare output folder (same structure, but files are overwritten after processing)
            string outputFolder = Path.Combine(workbookFolder, "Processed");
            Directory.CreateDirectory(outputFolder);

            // Supported Excel extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb" };
            string[] workbookFiles = Directory.GetFiles(workbookFolder, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;
            int failedCount = 0;

            foreach (string filePath in workbookFiles)
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // skip non‑Excel files

                Console.WriteLine($"Processing \"{Path.GetFileName(filePath)}\"...");

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Access the XmlMap collection
                    XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                    // Check if a map with the desired name already exists
                    bool mapExists = false;
                    foreach (XmlMap existingMap in xmlMaps)
                    {
                        if (string.Equals(existingMap.Name, desiredMapName, StringComparison.OrdinalIgnoreCase))
                        {
                            mapExists = true;
                            Console.WriteLine($"  XML map \"{desiredMapName}\" already exists. Skipping addition.");
                            break;
                        }
                    }

                    if (!mapExists)
                    {
                        // Add the XML map from the XSD file
                        int mapIndex = xmlMaps.Add(xmlSchemaPath);
                        XmlMap addedMap = xmlMaps[mapIndex];
                        addedMap.Name = desiredMapName; // set a friendly name
                        Console.WriteLine($"  Added XML map \"{desiredMapName}\" (Index: {mapIndex}).");
                    }

                    // Save the workbook (overwrite original or save to output folder)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                    Console.WriteLine($"  Saved processed workbook to \"{outputPath}\".");
                    processedCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Failed to process \"{Path.GetFileName(filePath)}\": {ex.Message}");
                    failedCount++;
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Batch processing completed. Success: {processedCount}, Failed: {failedCount}");
            Console.WriteLine($"Processed files are located in: {outputFolder}");
        }
    }
}
