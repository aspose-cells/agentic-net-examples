using System;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare an ArrayList containing GUID strings
        ArrayList guidList = new ArrayList();
        for (int i = 0; i < 5; i++)
        {
            guidList.Add(Guid.NewGuid().ToString()); // e.g., "3f2504e0-4f89-11d3-9a0c-0305e82c3301"
        }

        // Import the GUID list vertically starting at cell A1 (row 0, column 0)
        cells.ImportArrayList(guidList, 0, 0, true);

        // Convert each imported cell value to uppercase to ensure display in uppercase
        for (int row = 0; row < guidList.Count; row++)
        {
            Cell cell = cells[row, 0];
            if (cell.Value != null)
            {
                cell.PutValue(cell.StringValue.ToUpper());
            }
        }

        // Apply a text number format to keep the values as text (prevents Excel from auto‑formatting)
        Style textStyle = workbook.CreateStyle();
        textStyle.Number = 49; // Text format code
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        cells.CreateRange(0, 0, guidList.Count, 1).ApplyStyle(textStyle, flag);

        // Save the workbook
        workbook.Save("GuidListUppercase.xlsx");
    }
}