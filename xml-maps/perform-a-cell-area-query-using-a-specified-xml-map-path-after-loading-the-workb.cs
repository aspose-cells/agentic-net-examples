using System;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook that contains an XML map
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the workbook has at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Get the first XML map (or select a specific one as needed)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path to query
        string xmlPath = "/Root/Data/Item";

        // Query cell areas that are linked to the specified XML path
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Output the results
        if (cellAreas.Count > 0)
        {
            foreach (CellArea area in cellAreas)
            {
                Console.WriteLine($"Mapped area: Start({area.StartRow}, {area.StartColumn}) " +
                                  $"End({area.EndRow}, {area.EndColumn})");
                // Display the value of the top‑left cell in the area
                Console.WriteLine($"Value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
            }
        }
        else
        {
            Console.WriteLine("No cells are linked to the specified XML path.");
        }

        // Save the workbook (optional, as no modifications are made)
        workbook.Save("output.xlsx");
    }
}