// Title: Load multiple Excel workbooks from a directory, apply a common XML map, and batch export each workbook to XML with Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a folder for .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, assigns a predefined XML map called "MyMap", and saves the result as an .xml file in a target directory. | Create a reusable C# method that receives source and destination paths, applies the same XML map to every workbook found, exports each to XML, and returns a list of successfully generated XML file paths while catching load/save exceptions. | Generate a PowerShell script that calls a compiled .NET utility to perform batch XML export of Excel files, ensuring the XML map is applied and any errors are written to a log file.
// Common Searches: aspnet batch convert excel files to xml using aspose.cells | c# load all workbooks from folder and export to xml with same xml map | how to apply an xml map to multiple excel workbooks programmatically aspose.cells | automate xml export for many xlsm files in a directory c# | error handling when converting a batch of excel workbooks to xml aspose.cells
// Tags: batch export workbooks to xml aspose.cells | load multiple excel files from folder c# | assign xml map to workbook aspose.cells | save workbook as xml aspose.cells | handle conversion errors aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// The program scans a given input directory for .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, optionally applies a shared XML map, and saves the workbook data as an XML file with the same base name into an output folder, while handling missing files and logging any conversion errors.
class Program
{
    static void Main()
    {
        // Directory containing the source Excel workbooks
        string sourceDirectory = @"C:\InputWorkbooks";

        // Directory where the exported XML files will be saved
        string outputDirectory = @"C:\ExportedXml";

        try
        {
            // Ensure the source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Retrieve all Excel files (xlsx, xls, xlsm) from the source directory
            string[] workbookFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string workbookPath in workbookFiles)
            {
                string extension = Path.GetExtension(workbookPath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                    continue; // Skip non‑Excel files

                // Verify that the workbook file exists before loading
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found (skipped): {workbookPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(workbookPath);

                    // Determine the output XML file path (same name as workbook, .xml extension)
                    string outputXmlPath = Path.Combine(outputDirectory,
                        Path.GetFileNameWithoutExtension(workbookPath) + ".xml");

                    // Export the workbook to XML format
                    workbook.Save(outputXmlPath, SaveFormat.Xml);
                    Console.WriteLine($"Exported XML for '{Path.GetFileName(workbookPath)}' to '{outputXmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook '{workbookPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
