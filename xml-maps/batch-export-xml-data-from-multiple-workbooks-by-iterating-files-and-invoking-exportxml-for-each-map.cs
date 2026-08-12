// Title: Batch export XML maps from multiple Excel workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a directory of .xlsx files, loads each workbook with Aspose.Cells, detects every defined XML map, and writes each map to a uniquely named XML file in an output folder using Workbook.ExportXml.
// Keywords: Aspose.Cells | ExportXml | XML map export | batch Excel processing | C# .NET | multiple workbooks | folder iteration | save XML files | Windows | data integration
// Common Searches: aspocells export all xml maps from folder | c# batch export xml from excel workbooks | how to use Workbook.ExportXml for multiple files | export xml maps Aspose.Cells example | iterate excel files and export xml maps
// Developer Intent: Export every XML map present in each workbook of a folder to separate XML files.
// Use Cases: Automate generation of XML payloads for downstream services from a library of Excel templates. | Create per‑map XML exports for a reporting pipeline that consumes XML inputs. | Migrate Excel‑based XML data to an external system by saving each map as an individual file.
// AI Prompts: Write a reusable method that accepts input and output folder paths and uses Aspose.Cells to export all XML maps from each .xlsx file. | Add comprehensive error handling and logging to the batch XML export code to capture missing maps, file‑access errors, and permission issues. | Modify the batch export to export only XML maps whose names start with a given prefix. | Refactor the example into an async version that processes workbooks in parallel while preserving order of output files.

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that scans a directory of .xlsx files, loads each workbook with Aspose.Cells, detects every defined XML map, and writes each map to a uniquely named XML file in an output folder using Workbook.ExportXml.
class BatchXmlExport
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Folder where the exported XML files will be saved
        string outputFolder = @"C:\ExportedXml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (you can adjust the pattern if needed)
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string workbookPath in workbookFiles)
        {
            // Load the workbook (uses the standard Workbook(string) constructor)
            Workbook wb = new Workbook(workbookPath);

            // Check if the workbook has any XML maps defined
            if (wb.Worksheets.XmlMaps.Count > 0)
            {
                // Iterate through each XML map in the workbook
                for (int i = 0; i < wb.Worksheets.XmlMaps.Count; i++)
                {
                    XmlMap xmlMap = wb.Worksheets.XmlMaps[i];

                    // Construct a unique XML file name using the workbook name and map name
                    string workbookName = Path.GetFileNameWithoutExtension(workbookPath);
                    string xmlFileName = $"{workbookName}_{xmlMap.Name}.xml";
                    string xmlFullPath = Path.Combine(outputFolder, xmlFileName);

                    // Export the XML data for the current map (uses ExportXml(string, string))
                    wb.ExportXml(xmlMap.Name, xmlFullPath);
                }
            }
        }
    }
}
