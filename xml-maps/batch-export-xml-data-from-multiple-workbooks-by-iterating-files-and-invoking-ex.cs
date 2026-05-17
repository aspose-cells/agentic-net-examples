using System;
using System.IO;
using Aspose.Cells;

class BatchExportXml
{
    static void Main()
    {
        // Folder containing the source Excel workbooks
        string inputFolder = "InputWorkbooks";

        // Folder where the exported XML files will be saved
        string outputFolder = "ExportedXml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Iterate through all .xlsx files in the input folder
        foreach (string workbookPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Load the workbook from file
            Workbook wb = new Workbook(workbookPath);

            // Check if the workbook contains any XML maps
            if (wb.Worksheets.XmlMaps.Count > 0)
            {
                // Export each XML map to a separate file
                for (int i = 0; i < wb.Worksheets.XmlMaps.Count; i++)
                {
                    XmlMap map = wb.Worksheets.XmlMaps[i];

                    // Construct a unique output file name using workbook name and map name
                    string baseName = Path.GetFileNameWithoutExtension(workbookPath);
                    string outputPath = Path.Combine(outputFolder, $"{baseName}_{map.Name}.xml");

                    // Export the XML data for the current map
                    wb.ExportXml(map.Name, outputPath);
                }
            }
        }
    }
}