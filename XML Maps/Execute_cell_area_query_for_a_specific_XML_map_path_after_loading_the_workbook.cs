using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlMapQueryDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that contains an XML map
            string inputPath = "input.xlsx";

            // Load the workbook (uses workbook-load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (uses worksheet-access rule)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one XML map in the workbook
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Get the first XML map (uses worksheet-access rule for the collection)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XML element path you want to query
            string xmlElementPath = "/Root/Element";

            // Query cell areas that are linked to the specified XML path
            ArrayList cellAreas = worksheet.XmlMapQuery(xmlElementPath, xmlMap);

            // Output the results
            if (cellAreas.Count > 0)
            {
                Console.WriteLine($"Found {cellAreas.Count} cell area(s) linked to path '{xmlElementPath}':");
                foreach (CellArea area in cellAreas)
                {
                    // Display the start row/column of each area (zero‑based indices)
                    Console.WriteLine($"StartRow: {area.StartRow}, StartColumn: {area.StartColumn}");
                    // Optionally, display the cell value at the start of the area
                    Console.WriteLine($"Cell Value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
                }
            }
            else
            {
                Console.WriteLine($"No cells are linked to the XML path '{xmlElementPath}'.");
            }

            // (Optional) Save the workbook after the query if any changes were made
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}