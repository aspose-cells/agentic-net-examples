// Title: Use ImportArrayList to load ColorType enum numeric values and apply a custom number format that displays enum names in an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Import a C# ArrayList of ColorType enum integer values into a worksheet using Cells.ImportArrayList, then apply a style that maps each integer to its enum name via a custom number format. | Define a Style object, set its Custom property to a format string that translates numeric enum codes to their textual labels, and assign this style to the cells that were imported. | Write the workbook to disk so the resulting Excel file shows the enum names (Automatic, AutomaticIndex, RGB, IndexedColor, Theme) instead of the numeric codes.
// Common Searches: how to show enum names instead of numbers in an Excel file generated with Aspose.Cells C# | using ImportArrayList to bring enum values into a worksheet and display them as text | Aspose.Cells custom formatting to convert ColorType numeric codes to readable strings | C# example for mapping enum integer values to text in a saved XLSX with Aspose.Cells
// Tags: importarraylist enum values c# | enum display format aspnet | colorType enum to text aspose.cells | apply style to imported range aspose.cells | save workbook with enum names aspose.cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumDemo
{
    // Demonstrates importing the numeric values of the ColorType enum into a worksheet via Cells.ImportArrayList, creating a custom number format that maps each value to its enum name, applying the style to the imported range, and saving the workbook as EnumDisplayDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList containing the numeric values of the ColorType enum
            ArrayList enumValues = new ArrayList
            {
                (int)ColorType.Automatic,
                (int)ColorType.AutomaticIndex,
                (int)ColorType.RGB,
                (int)ColorType.IndexedColor,
                (int)ColorType.Theme
            };

            // Import the numeric enum values horizontally starting at cell A1
            // Parameters: (ArrayList, firstRow, firstColumn, isVertical)
            cells.ImportArrayList(enumValues, 0, 0, false);

            // Create a style with a custom number format that maps numbers to enum names
            Style enumStyle = workbook.CreateStyle();
            // Custom format: [=0]"Automatic";[=1]"AutomaticIndex";[=2]"RGB";[=3]"IndexedColor";[=4]"Theme";General
            enumStyle.Custom = "[=0]\"Automatic\";[=1]\"AutomaticIndex\";[=2]\"RGB\";[=3]\"IndexedColor\";[=4]\"Theme\";General";

            // Apply the style to the imported range (first row, columns A to E)
            for (int col = 0; col < enumValues.Count; col++)
            {
                cells[0, col].SetStyle(enumStyle);
            }

            // Save the workbook to a file
            workbook.Save("EnumDisplayDemo.xlsx");
        }
    }
}
