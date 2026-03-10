using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Add an XML map to the workbook using the XSD schema file
        int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

        // (Optional) Set a friendly name for the added XML map
        workbook.Worksheets.XmlMaps[mapIndex].Name = "MyXmlMap";

        // Save the workbook with the new XML map
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}