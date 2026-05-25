using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsEnumDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList with enum values (ColorType enum as an example)
            ArrayList enumValues = new ArrayList();
            enumValues.Add(ColorType.Automatic);       // 0
            enumValues.Add(ColorType.AutomaticIndex); // 1
            enumValues.Add(ColorType.RGB);            // 2
            enumValues.Add(ColorType.IndexedColor);   // 3
            enumValues.Add(ColorType.Theme);          // 4

            // Import the enum values vertically starting at cell A1 (row 0, column 0)
            cells.ImportArrayList(enumValues, 0, 0, true);

            // Create a style with a custom number format that maps the numeric values
            // to their corresponding enum names. The format uses conditional sections:
            // [=value]"text"; for each enum value, followed by a default General format.
            Style style = workbook.CreateStyle();
            style.Custom = "[=0]\"Automatic\";[=1]\"AutomaticIndex\";[=2]\"RGB\";[=3]\"IndexedColor\";[=4]\"Theme\";General";

            // Apply the style to the imported range (A1:A5)
            for (int row = 0; row < enumValues.Count; row++)
            {
                cells[row, 0].SetStyle(style);
            }

            // Save the workbook
            workbook.Save("EnumNamesDisplay.xlsx");
        }
    }
}