using System;
using System.IO;
using Aspose.Cells;

class BatchXmlMapAdder
{
    static void Main()
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\InputWorkbooks";

        // Folder where the modified workbooks will be saved
        string outputFolder = @"C:\OutputWorkbooks";

        // Path to the XML schema (XSD) that defines the XML map to be added
        string xmlMapPath = @"C:\Schema\MyMap.xsd";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files (you can adjust the pattern if needed)
        string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(filePath);

            // Add the XML map to the workbook's XmlMaps collection
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlMapPath);

            // Optionally set a friendly name for the map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = Path.GetFileNameWithoutExtension(xmlMapPath) + "_Map";

            // Save the modified workbook to the output folder (overwrites if exists)
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            workbook.Save(outputPath);

            // Release resources
            workbook.Dispose();
        }

        Console.WriteLine("Batch processing of workbooks completed successfully.");
    }
}