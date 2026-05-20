using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML data to import
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
                <Root>
                    <Data>
                        <Item>Value1</Item>
                        <Item>Value2</Item>
                    </Data>
                </Root>";

            // Convert XML string to a stream for ImportXml overload
            using (MemoryStream xmlStream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(xmlData);
                writer.Flush();
                xmlStream.Position = 0;

                // Import XML data into the first worksheet starting at cell A1
                workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the XML map that was created during import
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML map was created after import.");
                return;
            }

            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XML path we want to validate
            string xmlPath = "/Root/Data/Item";

            // Query cells linked to the specified XML path
            ArrayList linkedAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

            // Validate the query results
            if (linkedAreas.Count == 0)
            {
                Console.WriteLine($"No cells are linked to the XML path '{xmlPath}'.");
            }
            else
            {
                Console.WriteLine($"Cells linked to XML path '{xmlPath}':");
                foreach (CellArea area in linkedAreas)
                {
                    // For each linked area, output its address and current cell value
                    string address = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string value = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                    Console.WriteLine($" - {address}: \"{value}\"");
                }
            }

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("XmlMapValidationResult.xlsx");
        }
    }
}