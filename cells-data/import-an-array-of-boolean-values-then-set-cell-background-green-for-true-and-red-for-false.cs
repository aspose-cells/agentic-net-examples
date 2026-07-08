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
            bool[] boolValues = new bool[] { true, false, true, false, true };

            // Convert to object array for ImportObjectArray (Aspose.Cells does not have ImportArray for bool[])
            object[] objArray = Array.ConvertAll(boolValues, b => (object)b);

            // Import the boolean values vertically starting at cell A1 (row 0, column 0)
            cells.ImportObjectArray(objArray, 0, 0, true);

            // Define styles for true (green) and false (red) values
            Style trueStyle = workbook.CreateStyle();
            trueStyle.ForegroundColor = Color.LightGreen;
            trueStyle.Pattern = BackgroundType.Solid;

            Style falseStyle = workbook.CreateStyle();
            falseStyle.ForegroundColor = Color.LightCoral;
            falseStyle.Pattern = BackgroundType.Solid;

            // StyleFlag to apply cell shading
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true;

            // Apply background colors based on the boolean value in each cell
            for (int i = 0; i < boolValues.Length; i++)
            {
                Cell cell = cells[i, 0]; // column 0 (A), row i
                // Ensure the cell contains a boolean value
                if (cell.Type == CellValueType.IsBool)
                {
                    if (cell.BoolValue)
                    {
                        cell.SetStyle(trueStyle, flag);
                    }
                    else
                    {
                        cell.SetStyle(falseStyle, flag);
                    }
                }
            }

            // Save the workbook
            workbook.Save("BooleanArrayWithColors.xlsx");
        }
    }
}