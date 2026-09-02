// Title: Batch export XML maps from a folder of Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a directory for .xlsx files and uses Aspose.Cells Workbook.ExportXml to export a predefined list of XML maps from each workbook. | Modify the batch export program to read XML map names from an external JSON file and apply them to every workbook in the input folder. | Add comprehensive logging to the batch ExportXml script, recording the path of each processed workbook, the exported map name, and any errors encountered.
// Common Searches: aspnet c# batch export xml maps from multiple excel files using aspose.cells | how to loop through a folder of .xlsx files and call ExportXml for each map | c# program to extract specific XML maps from many workbooks with Aspose.Cells | read xml map names from json and export them with Aspose.Cells in bulk | error handling for ExportXml when processing a directory of workbooks
// Tags: Aspose.Cells ExportXml batch processing | C# iterate folder export XML maps | automated XML map extraction from Excel workbooks | read XML map list from JSON C# | logging ExportXml operations .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The program walks through all .xlsx files in a specified input directory, loads each workbook with Aspose.Cells, and exports the defined XML maps (e.g., Map1, Map2) to an output folder. Each exported file is named using the source workbook name and map identifier, with error handling and folder creation logic included.
class Program
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Folder where the exported XML files will be placed
        string outputFolder = @"C:\ExportedXml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // List of map names to export from each workbook
        List<string> mapNames = new List<string> { "Map1", "Map2" };

        // Iterate over all .xlsx files in the input folder
        foreach (string workbookPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Guard against missing files (should not happen with GetFiles, but added for safety)
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Export each specified map to an XML file
                foreach (string mapName in mapNames)
                {
                    // Construct a unique XML file name per workbook and map
                    string xmlFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_{mapName}.xml";
                    string xmlFullPath = Path.Combine(outputFolder, xmlFileName);

                    // Export the map to XML using the correct API
                    workbook.ExportXml(mapName, xmlFullPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
            }
        }
    }
}
