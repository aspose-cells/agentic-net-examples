using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsListBoxSelection
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate column A with sample items
            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Add a ListBox shape and bind it to the range A1:A5
            ListBox listBox = worksheet.Shapes.AddListBox(1, 0, 1, 0, 150, 100) as ListBox;
            listBox.InputRange = "A1:A5";
            listBox.SelectionType = SelectionType.Multi; // allow multiple selections

            // Select items whose values are present in column A
            // (In this example we simply select all items)
            int lastRow = worksheet.Cells.MaxDataRow; // last row with data in column A
            for (int row = 0; row <= lastRow; row++)
            {
                // Ensure the cell is not empty
                if (!string.IsNullOrEmpty(worksheet.Cells[row, 0].StringValue))
                {
                    // Select the item at the corresponding zero‑based index
                    listBox.SelectedItem(row, true);
                }
            }

            // Save the workbook
            workbook.Save("ListBoxSelectionBasedOnColumnA.xlsx");
        }
    }
}