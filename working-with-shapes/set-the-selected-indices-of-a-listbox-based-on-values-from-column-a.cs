// Title: Aspose.Cells .NET – Programmatically select ListBox items using column A indices
// Description: The C# sample builds a workbook, writes index numbers to cells A1‑A3, creates six list entries in column A, adds a ListBox shape bound to that range, enables multi‑selection, reads each index, checks it against the ListBox item count, marks the matching entries as selected, and saves the result as ListBoxSelectionBasedOnColumnA.xlsx.
// Keywords: Aspose.Cells | C# | ListBox shape | selected indices | InputRange | SelectionType.Multi | Excel automation | programmatic selection | read cell values | bind ListBox to range
// Common Searches: Aspose.Cells select ListBox items by index | C# set ListBox.SelectedItem from worksheet cells | bind ListBox to range Aspose.Cells example | read integer values from Excel to control ListBox | multi‑select ListBox using Aspose.Cells .NET
// Developer Intent: Programmatically pre‑select entries in a ListBox shape based on values stored in worksheet cells.
// Use Cases: Load a template where column A stores previously saved selection numbers and the ListBox reflects those choices on open. | Create a form‑like sheet where users enter selection codes in a column and the ListBox updates automatically to show the corresponding items. | Generate a report that highlights specific ListBox options after calculations write their indices to the worksheet.
// AI Prompts: Write C# code with Aspose.Cells that binds a ListBox to a named range and selects items according to integer values from another column, including out‑of‑range handling. | Show how to adapt the example to read a dynamic number of rows for selection indices and update the ListBox selection in one pass. | Explain how to add error logging when a cell contains a non‑numeric value or an index larger than the ListBox item count.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsListBoxSelectionDemo
{
    // The C# sample builds a workbook, writes index numbers to cells A1‑A3, creates six list entries in column A, adds a ListBox shape bound to that range, enables multi‑selection, reads each index, checks it against the ListBox item count, marks the matching entries as selected, and saves the result as ListBoxSelectionBasedOnColumnA.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data in column A (indices to be selected)
            // In a real scenario, this data could already exist in the worksheet.
            sheet.Cells["A1"].PutValue(0); // select first item
            sheet.Cells["A2"].PutValue(2); // select third item
            sheet.Cells["A3"].PutValue(4); // select fifth item

            // Populate list box items (also in column A for simplicity)
            for (int i = 0; i < 6; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Add a ListBox shape
            ListBox listBox = sheet.Shapes.AddListBox(1, 0, 1, 0, 150, 100) as ListBox;

            // Bind the list box to the range containing the items
            listBox.InputRange = "A1:A6";

            // Allow multiple selections
            listBox.SelectionType = SelectionType.Multi;

            // Read indices from column A (A1:A3 in this example) and select corresponding items
            for (int row = 0; row < 3; row++) // adjust range as needed
            {
                Cell cell = sheet.Cells[row, 0];
                if (cell.Value != null && int.TryParse(cell.Value.ToString(), out int index))
                {
                    // Ensure the index is within the item count
                    if (index >= 0 && index < listBox.ItemCount)
                    {
                        listBox.SelectedItem(index, true);
                    }
                }
            }

            // Save the workbook
            workbook.Save("ListBoxSelectionBasedOnColumnA.xlsx");
        }
    }
}
