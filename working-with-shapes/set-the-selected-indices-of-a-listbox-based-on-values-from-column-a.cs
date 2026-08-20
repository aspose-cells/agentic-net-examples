// Title: Programmatically set selected indices of a ListBox from column A with Aspose.Cells for .NET
// Description: Creates a workbook, fills cells A1:A10, adds a ListBox shape bound to that range, enables multi‑selection, and uses ListBox.SelectedItem to select items whose text meets a condition (e.g., contains “5” or “7”). The workbook is saved with the configured selections.
// Keywords: Aspose.Cells ListBox | C# ListBox SelectedItem | Excel shape ListBox programmatic selection | bind ListBox to range | multi‑select ListBox .NET | set ListBox selected indices | column A filter Aspose | Excel automation C#
// Common Searches: How to select items in an Aspose.Cells ListBox based on cell values | Set multiple selected indices of a ListBox shape using column A in C# | Aspose.Cells ListBox SelectedItem example | Programmatically pre‑select ListBox entries bound to a range | C# Aspose.Cells ListBox multi‑selection from column data
// Developer Intent: Select specific ListBox entries by evaluating the text in column A and applying programmatic selection.
// Use Cases: Automatically pre‑select ListBox items that contain certain keywords after populating column A with dynamic data. | Mark rows that meet business criteria (e.g., contain “5” or “7”) before exporting the workbook for reporting. | Enable multi‑selection on a ListBox bound to a range and reflect filter results through code.
// AI Prompts: Show C# code that reads column A, applies a custom filter, and uses ListBox.SelectedItem to select matching items in an Aspose.Cells workbook. | Give an example of retrieving and updating the selected indices of a ListBox shape after it is bound to an input range. | Explain how to safely handle empty or null cells when programmatically selecting ListBox items based on column values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, fills cells A1:A10, adds a ListBox shape bound to that range, enables multi‑selection, and uses ListBox.SelectedItem to select items whose text meets a condition (e.g., contains “5” or “7”). The workbook is saved with the configured selections.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with sample items (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].PutValue($"Item {i + 1}");
        }

        // Add a ListBox shape and bind it to the range A1:A10
        ListBox listBox = sheet.Shapes.AddListBox(1, 0, 1, 0, 150, 200) as ListBox;
        listBox.InputRange = "A1:A10";
        listBox.SelectionType = SelectionType.Multi; // allow multiple selections

        // Example logic: select items whose text contains "5" or "7"
        for (int i = 0; i < 10; i++)
        {
            string cellValue = sheet.Cells[i, 0].StringValue;
            if (cellValue.Contains("5") || cellValue.Contains("7"))
            {
                // Select the item at zero‑based index i
                listBox.SelectedItem(i, true);
            }
        }

        // Save the workbook with the configured ListBox
        workbook.Save("ListBoxSelectionFromColumnA.xlsx");
    }
}
