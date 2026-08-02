using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Sample XML content to be imported
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

        // Convert the XML string into a MemoryStream
        using (MemoryStream xmlStream = new MemoryStream())
        {
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(xmlData);
                writer.Flush();
                xmlStream.Position = 0; // Reset position to the beginning

                // Import XML data from the stream into the first worksheet at cell A1
                workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
            }
        }

        // Save the workbook with the imported linked cells
        workbook.Save("ImportedXml.xlsx");
    }
}