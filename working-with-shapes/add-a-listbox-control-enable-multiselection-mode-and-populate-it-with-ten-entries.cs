// Title: Add a Multi‑Select ListBox Shape to an Excel worksheet using Aspose.Cells for .NET
// Description: Creates a new workbook, fills cells A1:A10 with items, inserts a ListBox shape, links it to the range, enables multiple selections, optionally ties the result to cell B1, and saves the file as ListBoxMultiSelect.xlsx.
// Keywords: Aspose.Cells | .NET | C# | ListBox shape | multi‑select | AddListBox | InputRange | SelectionType.Multi | LinkedCell | Excel form control | populate range | interactive spreadsheet
// Common Searches: Aspose.Cells add ListBox with multi‑select | C# bind ListBox to cell range Aspose | Set linked cell for ListBox shape Aspose.Cells | Configure SelectionType.Multi in Aspose.Cells | Create Excel form controls with Aspose.Cells
// Developer Intent: Insert a ListBox control, connect it to a data range, allow multiple items to be chosen, and persist the workbook.
// Use Cases: Build an interactive report where users can choose several categories from a dropdown‑style ListBox. | Design a data‑entry form in Excel that captures multiple selections for downstream calculations. | Provide a dynamic option list populated from worksheet data, with selections reflected in another cell.
// AI Prompts: Generate C# code using Aspose.Cells to add a ListBox, set its InputRange to A1:A10, enable multi‑select, and link the output to B1. | Explain how to adjust the position and size of a ListBox shape after adding it with Aspose.Cells. | Show how to retrieve the values selected in a multi‑select ListBox when the workbook is opened in Excel.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, fills cells A1:A10 with items, inserts a ListBox shape, links it to the range, enables multiple selections, optionally ties the result to cell B1, and saves the file as ListBoxMultiSelect.xlsx.
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
        // Parameters: topRow, top, leftColumn, left, height, width
        ListBox listBox = sheet.Shapes.AddListBox(1, 0, 2, 0, 120, 100);

        // Bind the ListBox to the range containing the items
        listBox.InputRange = "A1:A10";

        // Enable multi‑selection mode
        listBox.SelectionType = SelectionType.Multi;

        // Optional: link the selected value(s) to a cell
        listBox.LinkedCell = "B1";

        // Save the workbook
        workbook.Save("ListBoxMultiSelect.xlsx");
    }
}
