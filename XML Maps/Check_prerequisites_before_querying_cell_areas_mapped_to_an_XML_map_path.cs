using System;
using System.Collections;
using Aspose.Cells;

class XmlMapQueryPrerequisiteDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Import sample XML to create an XML map in the workbook
        string xml = @"<?xml version='1.0' encoding='UTF-8'?>
            <Root>
                <Data>
                    <Item>Value1</Item>
                </Data>
            </Root>";
        workbook.ImportXml(xml, "Sheet1", 0, 0);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps are defined in the workbook. Query cannot be performed.");
            return;
        }

        // Retrieve the first XML map
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path to query
        string path = "/Root/Data/Item";

        // Query cell areas that are linked to the specified XML path
        ArrayList cellAreas = worksheet.XmlMapQuery(path, xmlMap);

        // Output the query results
        if (cellAreas.Count > 0)
        {
            foreach (CellArea area in cellAreas)
            {
                Console.WriteLine($"Mapped cell found at Row {area.StartRow + 1}, Column {area.StartColumn + 1}");
                Console.WriteLine($"Cell value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
            }
        }
        else
        {
            Console.WriteLine("No cells are mapped to the specified XML path.");
        }

        // Save the workbook (optional)
        workbook.Save("XmlMapQueryDemo.xlsx", SaveFormat.Xlsx);
    }
}