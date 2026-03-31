using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook file
            Workbook workbook = new Workbook("input.xlsx");

            // Add an XML map to the workbook using an XSD or XML schema file
            // The Add method returns the index of the newly added map
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

            // Retrieve the XmlMap object (optional, e.g., to set a friendly name)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyDataMap";

            // Import XML data into the first worksheet starting at cell A1 (row 0, column 0)
            // This binds the XML data to the workbook using the previously added map
            workbook.ImportXml("data.xml", workbook.Worksheets[0].Name, 0, 0);

            // Save the workbook with the XML map and bound data
            workbook.Save("output.xlsx");
        }
    }
}