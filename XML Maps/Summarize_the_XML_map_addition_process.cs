using System;
using Aspose.Cells;

namespace XmlMapAdditionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: workbook-create)
            Workbook workbook = new Workbook();

            // Access the collection of XML maps in the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Add an XML map by providing the path or URL to an XSD or XML file
            // The Add method returns the index of the newly added map
            int mapIndex = xmlMaps.Add("schema.xsd");

            // Retrieve the added XmlMap using the returned index
            XmlMap xmlMap = xmlMaps[mapIndex];

            // Optionally set a friendly name for the map
            xmlMap.Name = "SampleMap";

            // Example: link a cell to a specific element in the XML map
            // This demonstrates how the map can be used after addition
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Element");

            // Save the workbook (lifecycle rule: workbook-save)
            workbook.Save("XmlMapDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}