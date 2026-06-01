using System;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the XML file that contains both the schema and the data.
            // The file should exist in the same directory as the executable.
            string xmlPath = "data.xml";

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Add an XML map to the workbook using the XML file.
            // The Add method returns the index of the newly added map.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DataMap";

            // Import the XML data into the first worksheet starting at cell A1.
            // This populates the worksheet with the XML content.
            workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Get the first worksheet and its cells collection.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Link specific cells to elements in the XML map.
            // Adjust the row, column, and XML path as needed for your XML structure.
            // Example: link cell A1 to the XML element /Root/Title
            cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Title");

            // Example: link cell B2 to the XML element /Root/Details/Price
            cells.LinkToXmlMap(xmlMap.Name, 1, 1, "/Root/Details/Price");

            // Save the workbook with the linked XML map.
            workbook.Save("output.xlsx");
        }
    }
}