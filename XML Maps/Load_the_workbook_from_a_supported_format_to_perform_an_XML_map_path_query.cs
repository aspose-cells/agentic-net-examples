using System;
using System.Collections;
using Aspose.Cells;

class XmlMapPathQueryDemo
{
    static void Main()
    {
        // Path to the XML file (or an Excel file containing XML map)
        string inputPath = "input.xml";

        // Configure XML load options to enable XML mapping
        XmlLoadOptions loadOptions = new XmlLoadOptions();
        loadOptions.IsXmlMap = true;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one XML map in the workbook
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Retrieve the first XML map
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path to query
        string xmlPath = "/Root/Item";

        // Query cell areas linked to the specified XML path
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Output the query results
        Console.WriteLine($"Found {cellAreas.Count} cell area(s) for path '{xmlPath}':");
        foreach (CellArea area in cellAreas)
        {
            // Get the first cell in the area
            Cell cell = worksheet.Cells[area.StartRow, area.StartColumn];
            Console.WriteLine($"Row: {area.StartRow + 1}, Column: {area.StartColumn + 1}, Value: {cell.StringValue}");
        }
    }
}