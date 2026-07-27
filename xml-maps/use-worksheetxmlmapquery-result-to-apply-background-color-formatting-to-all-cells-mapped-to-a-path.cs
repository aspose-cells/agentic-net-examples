using System;
using System.Collections;
using System.Drawing;
using Aspose.Cells;

class ApplyXmlMapBackgroundColor
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps found in the workbook.");
            return;
        }

        // Use the first XML map (adjust index if needed)
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML element path whose linked cells should be formatted
        string xmlPath = "/Root/Item"; // <-- change to your actual path

        // Query the worksheet for cell areas linked to the specified XML path
        ArrayList cellAreas = sheet.XmlMapQuery(xmlPath, xmlMap);

        // Create a style with the desired background color
        Style bgStyle = workbook.CreateStyle();
        bgStyle.ForegroundColor = Color.Yellow;          // background color
        bgStyle.Pattern = BackgroundType.Solid;          // solid fill

        // Apply the style to every cell within each returned CellArea
        foreach (CellArea area in cellAreas)
        {
            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.SetStyle(bgStyle);
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}