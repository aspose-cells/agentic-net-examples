using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlImportLinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Sample XML data
            string xmlData = @"<Products>
                                    <Product>
                                        <Name>Laptop</Name>
                                        <Price>999.99</Price>
                                    </Product>
                                    <Product>
                                        <Name>Phone</Name>
                                        <Price>699.99</Price>
                                    </Product>
                               </Products>";

            // Convert XML string to a MemoryStream
            using (MemoryStream xmlStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(xmlData);
                    writer.Flush();
                    xmlStream.Position = 0;

                    // Import XML data into the first worksheet starting at cell A1
                    wb.ImportXml(xmlStream, "Sheet1", 0, 0);
                }
            }

            // Retrieve the name of the automatically created XML map
            string mapName = wb.Worksheets.XmlMaps[0].Name;

            // Get the first worksheet and its cells collection
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Link specific cells to XML map elements
            // Link cell A1 to the first product name
            cells.LinkToXmlMap(mapName, 0, 0, "/Products/Product[1]/Name");
            // Link cell B1 to the first product price
            cells.LinkToXmlMap(mapName, 0, 1, "/Products/Product[1]/Price");
            // Link cell A2 to the second product name
            cells.LinkToXmlMap(mapName, 1, 0, "/Products/Product[2]/Name");
            // Link cell B2 to the second product price
            cells.LinkToXmlMap(mapName, 1, 1, "/Products/Product[2]/Price");

            // Save the workbook with the linked XML map
            wb.Save("XmlImportLinked.xlsx");
        }
    }
}