// Title: C# program to batch export XML maps from multiple Excel workbooks using Aspose.Cells .NET
// Common Searches: aspocells export xmlmap from all workbooks in a folder c# | batch export xml maps for each Excel file using Aspose.Cells .NET | iterate through .xlsx files and save each XmlMap as separate XML in C# | how to export multiple XmlMap definitions from workbooks with Aspose.Cells | c# script to export xml maps from a directory of Excel files
// Tags: batch export xml maps Aspose.Cells .NET | export XmlMap from Excel workbook C# | process multiple Excel files Aspose.Cells | safe filename generation for XmlMap export | supported Excel extensions Aspose.Cells batch

using System;
using System.IO;
using Aspose.Cells;

namespace BatchXmlExport
{
    // The sample validates input and optional output folders, creates an export directory, scans the specified folder for supported Excel formats, loads each workbook with Aspose.Cells, checks for defined XmlMaps, and exports each map to a uniquely named XML file using a safe‑filename routine. Progress is logged and load or export errors are handled gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: BatchXmlExport <InputFolderPath> [OutputFolderPath]");
                return;
            }

            string inputFolder = args[0];
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Error: Input folder does not exist: {inputFolder}");
                return;
            }

            string outputFolder = args.Length > 1 ? args[1] : Path.Combine(inputFolder, "ExportedXml");
            try
            {
                Directory.CreateDirectory(outputFolder);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Unable to create output folder '{outputFolder}'. {ex.Message}");
                return;
            }

            // Supported Excel extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods" };
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            int processedCount = 0;

            foreach (string filePath in workbookFiles)
            {
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLowerInvariant()) < 0)
                    continue; // Skip non‑Excel files

                Console.WriteLine($"Processing workbook: {Path.GetFileName(filePath)}");
                try
                {
                    // Load the workbook
                    using (Workbook workbook = new Workbook(filePath))
                    {
                        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
                        if (xmlMaps.Count == 0)
                        {
                            Console.WriteLine("  No XmlMap defined in this workbook.");
                            continue;
                        }

                        // Export each XmlMap
                        for (int i = 0; i < xmlMaps.Count; i++)
                        {
                            XmlMap xmlMap = xmlMaps[i];
                            string safeMapName = MakeFileSystemSafe(xmlMap.Name);
                            string outputFileName = $"{Path.GetFileNameWithoutExtension(filePath)}_{safeMapName}.xml";
                            string outputPath = Path.Combine(outputFolder, outputFileName);

                            try
                            {
                                workbook.ExportXml(xmlMap.Name, outputPath);
                                Console.WriteLine($"  Exported XmlMap '{xmlMap.Name}' to '{outputFileName}'.");
                            }
                            catch (Exception exportEx)
                            {
                                Console.WriteLine($"  Error exporting XmlMap '{xmlMap.Name}': {exportEx.Message}");
                            }
                        }
                    }

                    processedCount++;
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"  Failed to load workbook '{Path.GetFileName(filePath)}': {loadEx.Message}");
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Batch export completed. Processed {processedCount} workbook(s).");
            Console.WriteLine($"Exported XML files are located in: {outputFolder}");
        }

        /// <summary>
        /// Replaces characters that are invalid in file names with an underscore.
        /// </summary>
        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }
    }
}
