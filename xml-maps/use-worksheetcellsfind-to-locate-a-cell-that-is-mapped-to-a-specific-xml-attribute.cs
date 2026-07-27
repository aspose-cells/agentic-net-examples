using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // for FindOptions enums

namespace AsposeCellsXmlFindDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Sample XML containing an attribute "id".
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item id='123'>Value1</Item>
    <Item id='456'>Value2</Item>
</Root>";

                // 3. Import the XML into the worksheet.
                //    This creates an XML map automatically (default name is "Map1").
                workbook.ImportXml(xml, sheet.Name, 0, 0);

                // 4. Use the default map name created by ImportXml.
                const string mapName = "Map1";

                // 5. Link cell A1 (row 0, column 0) to the first item's "id" attribute.
                //    Path syntax: "/Root/Item/@id"
                cells.LinkToXmlMap(mapName, 0, 0, "/Root/Item/@id");

                // 6. Retrieve and display the linked cell's value (populated from XML).
                Cell linkedCell = cells[0, 0];
                Console.WriteLine($"Linked cell initial value: {linkedCell.StringValue}");

                // 7. Use Find to locate the cell that holds the attribute value "123".
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values,
                    LookAtType = LookAtType.EntireContent,
                    CaseSensitive = false
                };

                Cell foundCell = cells.Find("123", null, findOptions);
                if (foundCell != null)
                {
                    Console.WriteLine($"Found cell mapped to attribute 'id' at: {foundCell.Name}");
                }
                else
                {
                    Console.WriteLine("Attribute value not found.");
                }

                // 8. Save the workbook (optional).
                string outputPath = "XmlAttributeFindDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {System.IO.Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}