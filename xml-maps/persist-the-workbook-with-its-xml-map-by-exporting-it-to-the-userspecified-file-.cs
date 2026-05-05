using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook (replace with actual path as needed)
        string sourcePath = "Book1.xlsx";

        // Path where the workbook (with XML map) will be saved
        string destinationPath = "WorkbookWithMap.xlsx";

        // Path for the separate exported XML file (optional, demonstrates ExportXml)
        string exportedXmlPath = "ExportedData.xml";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(sourcePath);

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }

        // Retrieve the first XML map (you can select a different one by index or name)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Export the XML data linked to the map to a separate .xml file
        workbook.ExportXml(xmlMap.Name, exportedXmlPath);
        Console.WriteLine($"XML data exported to '{exportedXmlPath}'.");

        // Prepare save options that embed the XML map into the saved workbook
        XmlSaveOptions saveOptions = new XmlSaveOptions
        {
            // Specify the name of the XML map to be included in the saved file
            XmlMapName = xmlMap.Name
        };

        // Save the workbook to the desired format (XLSX in this example) with the XML map
        workbook.Save(destinationPath, saveOptions);
        Console.WriteLine($"Workbook saved with XML map to '{destinationPath}'.");
    }
}