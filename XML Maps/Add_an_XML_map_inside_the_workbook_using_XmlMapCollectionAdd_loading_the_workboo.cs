using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx"); // workbook-load

        // Add an XML map to the workbook using a schema or XML file path
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd"); // XmlMapCollection.Add

        // Optionally set a friendly name for the map
        workbook.Worksheets.XmlMaps[mapIndex].Name = "MyXmlMap";

        // Save the workbook with the new XML map
        workbook.Save("output.xlsx", SaveFormat.Xlsx); // workbook-save
    }
}