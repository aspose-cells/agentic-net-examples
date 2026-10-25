// Title: How to batch export XML maps from multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export all XML maps from each workbook in a folder | iterate over .xlsx files and save each XmlMap to a separate XML file using Aspose.Cells | batch process Excel workbooks to generate XML files for every defined XML map in .NET | automate exporting XML maps from many Excel files with Aspose.Cells C# | Aspose.Cells ExportXml for multiple workbooks in a directory
// Tags: batch export xml maps Aspose.Cells C# | iterate excel workbooks export xml Aspose.Cells | export xml map per workbook Aspose.Cells | process multiple .xlsx files Aspose.Cells | save xml map with workbook name C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchExport
{
    // The example scans a given input directory for .xlsx files, loads each workbook with Aspose.Cells, loops through all defined XmlMaps, and exports each map to an individual XML file named with the workbook and map identifiers. Exported files are written to a specified output folder that is created if missing, and errors are logged while processing continues.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel workbooks
            string inputFolder = @"C:\InputWorkbooks";
            // Output folder where exported XML files will be saved
            string outputFolder = @"C:\ExportedXml";

            BatchExportXml(inputFolder, outputFolder);
        }

        /// <param name="inputFolder">Folder containing the source workbooks.</param>
        /// <param name="outputFolder">Folder where the XML files will be written.</param>
        static void BatchExportXml(string inputFolder, string outputFolder)
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all Excel files (adjust the pattern as needed)
            string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string workbookPath in workbookFiles)
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    continue;
                }

                try
                {
                    // Load the workbook using Aspose.Cells
                    using (Workbook workbook = new Workbook(workbookPath))
                    {
                        // Retrieve the workbook name without extension for naming exported files
                        string workbookName = Path.GetFileNameWithoutExtension(workbookPath);

                        // Iterate through each XML map defined in the workbook
                        foreach (XmlMap xmlMap in workbook.Worksheets.XmlMaps)
                        {
                            // Construct a unique file name for the exported XML
                            string xmlFileName = $"{workbookName}_{xmlMap.Name}.xml";
                            string xmlOutputPath = Path.Combine(outputFolder, xmlFileName);

                            // Export the XML map to the specified file
                            // Using the overload that accepts the map name and output path
                            workbook.ExportXml(xmlMap.Name, xmlOutputPath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other workbooks
                    Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                }
            }
        }
    }
}
