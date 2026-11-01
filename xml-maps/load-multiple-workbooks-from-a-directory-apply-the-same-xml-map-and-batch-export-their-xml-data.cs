// Title: Batch export XML files from all .xlsx workbooks in a directory using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates every .xlsx file in a specified folder, loads each workbook with Aspose.Cells, and saves it as an XML file to a target directory. | Add comprehensive try‑catch blocks to the batch conversion script so that failures for individual workbooks are logged but do not stop processing of remaining files. | Modify the batch export program to create the output folder if it does not exist and to use the workbook name as the XML file name.
// Common Searches: c# Aspose.Cells export all Excel files in a folder to XML | how to loop through .xlsx files and save as .xml using Aspose.Cells .NET | batch convert multiple workbooks to XML with error handling Aspose.Cells | save each workbook as XML in a separate output directory using Aspose.Cells | Aspose.Cells SaveFormat.Xml example for bulk processing
// Tags: Aspose.Cells batch workbook to XML conversion | C# enumerate .xlsx files and export XML | SaveFormat.Xml usage in bulk Excel processing | directory based Excel to XML conversion .NET | error resilient workbook export with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program scans a given input folder for .xlsx files, loads each workbook with Aspose.Cells, and saves its content as an XML file (using the same base name) into a specified output folder, creating the output directory if needed and handling individual file errors without aborting the batch.
class Program
{
    static void Main()
    {
        // Directory containing the source Excel workbooks
        string excelDirectory = @"C:\InputExcels";

        // Directory where the exported XML files will be saved
        string exportDirectory = @"C:\ExportedXml";

        try
        {
            // Ensure the export directory exists
            Directory.CreateDirectory(exportDirectory);

            // Validate the input directory
            if (!Directory.Exists(excelDirectory))
                throw new DirectoryNotFoundException($"Input directory not found: {excelDirectory}");

            // Retrieve all .xlsx files from the specified directory
            string[] workbookFiles = Directory.GetFiles(excelDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string workbookPath in workbookFiles)
            {
                try
                {
                    // Verify the workbook file exists before loading
                    if (!File.Exists(workbookPath))
                    {
                        Console.WriteLine($"Workbook file not found: {workbookPath}");
                        continue;
                    }

                    // Load the workbook from file
                    Workbook workbook = new Workbook(workbookPath);

                    // Determine the output XML file name (same as workbook name but with .xml extension)
                    string xmlFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".xml";
                    string xmlOutputPath = Path.Combine(exportDirectory, xmlFileName);

                    // Export the workbook's data to XML format
                    workbook.Save(xmlOutputPath, SaveFormat.Xml);
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
