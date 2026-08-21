// Title: C# – Add a Multi‑Select ListBox Shape Bound to a Cell Range Using Aspose.Cells
// Description: Creates a new workbook, writes ten items to A1:A10, inserts a ListBox shape at row 1 column 2, binds it to the range, enables multi‑select mode, and saves the file as ListBoxMultiSelection.xlsx.
// Keywords: Aspose.Cells ListBox | C# multi‑select ListBox | Excel form control binding | Add ListBox shape programmatically | InputRange property Aspose
// Common Searches: aspocells add listbox multi select | bind listbox to cell range c# | enable multi‑selection listbox aspocells | listbox shape example aspocells .net
// Developer Intent: Insert a ListBox control, link it to cells A1:A10, and turn on multi‑selection.
// Use Cases: Provide end‑users with a selectable list of categories pre‑filled from worksheet data. | Create a template where a multi‑select ListBox filters report rows based on chosen items. | Design a data‑entry form that lets users pick several options from a dynamically populated list.
// AI Prompts: Generate C# code that adds a multi‑select ListBox to an Excel sheet with Aspose.Cells and binds it to a dynamic range. | Show how to read the selected values from a multi‑select ListBox after opening the workbook with Aspose.Cells. | Explain how to customize font, colors, and dimensions of a ListBox shape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes ten items to A1:A10, inserts a ListBox shape at row 1 column 2, binds it to the range, enables multi‑select mode, and saves the file as ListBoxMultiSelection.xlsx.
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
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
        ListBox listBox = sheet.Shapes.AddListBox(1, 0, 2, 0, 120, 200);

        // Bind the ListBox to the range containing the ten items
        listBox.InputRange = "A1:A10";

        // Enable multi‑selection mode
        listBox.SelectionType = SelectionType.Multi;

        // Save the workbook
        workbook.Save("ListBoxMultiSelection.xlsx");
    }
}
