using System;
using System.Collections;
using Aspose.Cells;

class RetrieveXmlMapPaths
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps are present in the workbook.");
            return;
        }

        // Get the first XML map (or select a specific one by index/name)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path you want to query
        // Adjust this path to match the schema of your XML map
        string xmlPath = "/Root/Data/Item";

        // Query the worksheet for cell areas linked to the specified XML path
        ArrayList cellAreas = worksheet.XmlMapQuery(xmlPath, xmlMap);

        // Output the results
        if (cellAreas.Count > 0)
        {
            foreach (CellArea area in cellAreas)
            {
                Console.WriteLine($"Mapped area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, EndRow={area.EndRow}, EndColumn={area.EndColumn}");
                // Example: display the value of the first cell in the area
                Console.WriteLine($"Cell value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
            }
        }
        else
        {
            Console.WriteLine($"No cells are mapped to the XML path '{xmlPath}'.");
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx");
    }
}