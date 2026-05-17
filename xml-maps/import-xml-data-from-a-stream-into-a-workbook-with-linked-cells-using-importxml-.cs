using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Sample XML data to be imported
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

        // Convert the XML string to a readable MemoryStream
        byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
        using (MemoryStream xmlStream = new MemoryStream(xmlBytes))
        {
            // Import the XML data into the first worksheet starting at cell A1
            workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
        }

        // After import, an XML map is created. Link a cell to a specific XML element.
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];
            // Link cell C1 (row 0, column 2) to the Name element of the first Product
            workbook.Worksheets[0].Cells.LinkToXmlMap(xmlMap.Name, 0, 2, "/Products/Product[1]/Name");
        }

        // Save the workbook with the imported XML and linked cell
        workbook.Save("ImportXmlLinked.xlsx");
    }
}