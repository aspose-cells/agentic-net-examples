using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate ten entries in column A (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue($"Item {i + 1}");
        }

        // Add a ListBox to the worksheet
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        ListBox listBox = sheet.Shapes.AddListBox(2, 0, 2, 0, 120, 200);

        // Bind the ListBox to the range containing the ten items
        listBox.InputRange = "A1:A10";

        // Enable multi‑selection mode
        listBox.SelectionType = SelectionType.Multi;

        // Optional: link the selected value to a cell
        listBox.LinkedCell = "B1";

        // Save the workbook
        workbook.Save("ListBoxMultiSelect.xlsx");
    }
}