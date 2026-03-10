using System;
using System.Collections;
using Aspose.Cells;

class XmlMapQueryDemo
{
    static void Main()
    {
        // Load an existing workbook (workbook-load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (worksheet-access rule)
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Retrieve the first XML map
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path to query
        string xmlPath = "/Root/Item";

        // Query cell areas that are linked to the specified XML path (Worksheet.XmlMapQuery)
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Output the results
        if (cellAreas.Count > 0)
        {
            foreach (CellArea area in cellAreas)
            {
                // For each returned area, display the start cell address and its value
                int row = area.StartRow;
                int column = area.StartColumn;
                string cellName = CellsHelper.CellIndexToName(row, column);
                string cellValue = worksheet.Cells[row, column].StringValue;
                Console.WriteLine($"Cell {cellName} maps to XML path '{xmlPath}'. Value: {cellValue}");
            }
        }
        else
        {
            Console.WriteLine($"No cells are mapped to the XML path '{xmlPath}'.");
        }

        // Save the workbook if any modifications are needed (workbook-save rule)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}