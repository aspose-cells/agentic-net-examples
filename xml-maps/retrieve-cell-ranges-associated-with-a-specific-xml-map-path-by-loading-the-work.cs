using System;
using System.Collections;
using Aspose.Cells;

class RetrieveXmlMapRanges
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or use the name)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps are defined in the workbook.");
            return;
        }

        // Get the XML map to query (using the first map in this example)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path you want to query
        // Adjust this path to match the structure of your XML map
        string xmlPath = "/Root/Item";

        // Query the worksheet for cell areas linked to the specified XML path
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Output the results
        if (cellAreas.Count > 0)
        {
            Console.WriteLine($"Found {cellAreas.Count} cell area(s) for XML path '{xmlPath}':");
            foreach (CellArea area in cellAreas)
            {
                // Display the start and end coordinates of each area
                Console.WriteLine($"Area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                                  $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");

                // Retrieve and display the value of the first cell in the area
                string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                Console.WriteLine($"Value at start cell ({area.StartRow}, {area.StartColumn}): {cellValue}");
            }
        }
        else
        {
            Console.WriteLine($"No cells are linked to the XML path '{xmlPath}'.");
        }

        // Save the workbook (optional, as no changes are made)
        workbook.Save("output.xlsx");
    }
}