using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsBooleanImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Boolean array to import
            bool[] boolValues = new bool[] { true, false, true, true, false };

            // Convert to object array because ImportObjectArray expects object[]
            object[] objArray = new object[boolValues.Length];
            for (int i = 0; i < boolValues.Length; i++)
                objArray[i] = boolValues[i];

            // Import the boolean values vertically starting at cell A1 (row 0, column 0)
            // Using the rule: Cells.ImportObjectArray(object[], int, int, bool)
            cells.ImportObjectArray(objArray, 0, 0, true);

            // Apply background colors: green for true, red for false
            for (int i = 0; i < boolValues.Length; i++)
            {
                Cell cell = cells[i, 0];

                // Create a new style for each cell
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = boolValues[i] ? Color.Green : Color.Red;

                // Apply the style to the cell
                cell.SetStyle(style);
            }

            // Save the workbook
            workbook.Save("BooleanImportWithColors.xlsx");
        }
    }
}