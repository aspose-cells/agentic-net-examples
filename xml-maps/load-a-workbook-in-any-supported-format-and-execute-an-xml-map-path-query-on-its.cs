using System;
using System.Collections;
using Aspose.Cells;

class XmlMapQueryDemo
{
    static void Main()
    {
        // Load a workbook (replace with your actual file path and format)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps are present in the workbook.");
            return;
        }

        // Retrieve the first XML map
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path you want to query
        // Adjust this path to match the schema of your XML map
        string xmlPath = "/Root/Data/Item";

        // Execute the XML map query on the worksheet
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Process and display the query results
        if (cellAreas.Count == 0)
        {
            Console.WriteLine("No cells are mapped to the specified XML path.");
        }
        else
        {
            foreach (CellArea area in cellAreas)
            {
                // Convert cell indices to A1 style address
                string address = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                // Retrieve the cell value as a string
                string value = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                Console.WriteLine($"Mapped cell: {address}, Value: {value}");
            }
        }
    }
}