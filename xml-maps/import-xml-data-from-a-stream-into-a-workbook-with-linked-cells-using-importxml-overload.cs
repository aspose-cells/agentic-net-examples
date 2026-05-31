using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class ImportXmlFromStreamDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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

                    // Import the XML data from the stream into the first worksheet at cell A1 (row 0, column 0)
                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                }
            }

            // Demonstrate linking a cell to the imported XML map (optional)
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                string mapName = workbook.Worksheets.XmlMaps[0].Name;

                // Link cell A1 to the first product's Name element in the XML map
                workbook.Worksheets[0].Cells.LinkToXmlMap(mapName, 0, 0, "/Products/Product[1]/Name");
            }

            // Save the workbook to a file
            string outputPath = "ImportedFromStream.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}