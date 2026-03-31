using System;
using Aspose.Cells;

class XmlMapAssociationDemo
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Retrieve the first XML map from the collection
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Associate specific cells with elements in the XML map
        // Link cell A1 (row 0, column 0) to the XML element "/Root/Item1"
        cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item1");

        // Link cell B2 (row 1, column 1) to the XML element "/Root/Item2"
        cells.LinkToXmlMap(xmlMap.Name, 1, 1, "/Root/Item2");

        // Save the workbook with the new XML map associations
        workbook.Save("output.xlsx");
    }
}