using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryValidationDemo
    {
        public static void Run()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample XML data to create an XML map
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item>Value1</Item>
    <Item>Value2</Item>
</Root>";

            // Import the XML into the worksheet to generate an XML map (creation rule)
            workbook.ImportXml(xmlData, worksheet.Name, 0, 0);

            // Retrieve the generated XML map
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps were created.");
                return;
            }
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XML path we want to query
            string xmlPath = "/Root/Item";

            // ----- Validation of required conditions before querying -----
            // 1. Ensure the worksheet reference is valid
            if (worksheet == null)
            {
                Console.WriteLine("Worksheet is null.");
                return;
            }

            // 2. Ensure the XML map reference is valid
            if (xmlMap == null)
            {
                Console.WriteLine("XML map is null.");
                return;
            }

            // 3. Ensure the path string is not null or empty
            if (string.IsNullOrWhiteSpace(xmlPath))
            {
                Console.WriteLine("XML path is empty.");
                return;
            }

            // Perform the query to get cell areas linked to the specified XML path
            ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

            // Output the results
            if (cellAreas.Count > 0)
            {
                foreach (CellArea area in cellAreas)
                {
                    Console.WriteLine($"Mapped area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                                      $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");
                    // Example: read the first cell's value in the area
                    string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                    Console.WriteLine($"First cell value in area: {cellValue}");
                }
            }
            else
            {
                Console.WriteLine("No cells are mapped to the specified XML path.");
            }

            // Save the workbook (save rule)
            workbook.Save("XmlMapQueryDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XmlMapQueryValidationDemo.Run();
        }
    }
}