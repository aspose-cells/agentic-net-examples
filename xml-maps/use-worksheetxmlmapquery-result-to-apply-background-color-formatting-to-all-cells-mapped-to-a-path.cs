using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapFormatting
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample XML data (used to create an XML map)
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Data>
        <Item>Value1</Item>
        <Item>Value2</Item>
    </Data>
</Root>";

                // Import the XML into the worksheet – this also creates an XML map
                workbook.ImportXml(xml, sheet.Name, 0, 0);

                // Retrieve the created XML map (first map in the collection)
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Link some cells to the XML map (for demonstration)
                // Link cell A1 to the first Item element
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Data/Item[1]");
                // Link cell A2 to the second Item element
                cells.LinkToXmlMap(xmlMap.Name, 1, 0, "/Root/Data/Item[2]");

                // Define the XML path we want to format cells for
                string path = "/Root/Data/Item";

                // Query the worksheet for all cell areas mapped to the specified path
                ArrayList cellAreas = sheet.XmlMapQuery(path, xmlMap);

                // Prepare a style with a solid background color (light yellow)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightYellow;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to every cell within each returned CellArea
                foreach (CellArea area in cellAreas)
                {
                    int rowCount = area.EndRow - area.StartRow + 1;
                    int colCount = area.EndColumn - area.StartColumn + 1;

                    // Create a range covering the current CellArea
                    Aspose.Cells.Range range = cells.CreateRange(area.StartRow, area.StartColumn, rowCount, colCount);

                    // Apply the background style to the entire range
                    range.SetStyle(style);
                }

                // Define output file name
                string outputPath = "XmlMapFormatted.xlsx";

                // Save the workbook with the applied formatting
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}