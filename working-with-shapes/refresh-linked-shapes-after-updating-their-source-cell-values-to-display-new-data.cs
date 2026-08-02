// Title: Refresh a linked ListBox shape after changing its linked cell using Aspose.Cells for .NET
// Description: Demonstrates how to synchronize a ListBox (or other form control) with its linked cell after the cell value changes. The example creates a workbook, adds a ListBox linked to a range, sets the linked cell, calls UpdateSelectedValue on the shape and on the worksheet's Shapes collection, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | linked shape refresh | UpdateSelectedValue | ListBox linked cell | form control synchronization | worksheet.Shapes.UpdateSelectedValue | Excel shape refresh programmatically
// Common Searches: Aspose.Cells refresh linked shape after cell update | UpdateSelectedValue for ListBox in Aspose.Cells | How to sync ListBox selection with linked cell .NET | Refresh all form controls in worksheet Aspose.Cells | C# Aspose.Cells linked cell shape update
// Developer Intent: Synchronize linked form controls with their source cells so the visual selection reflects the latest cell values.
// Use Cases: After programmatically changing the value of a linked cell (e.g., B1), call sheet.Shapes.UpdateSelectedValue() to refresh the ListBox selection. | When multiple controls (ListBox, ComboBox, CheckBox) are linked to cells, invoke UpdateSelectedValue on each shape or on the worksheet's Shapes collection to keep all controls in sync after bulk data modifications. | Refresh linked shapes before saving a workbook to ensure the on‑screen representation matches the underlying data.
// AI Prompts: Generate C# code that updates several linked ListBox and ComboBox shapes after modifying their linked cells using Aspose.Cells. | Explain the difference between Shape.UpdateSelectedValue() and Worksheet.Shapes.UpdateSelectedValue() and provide scenarios for each. | Provide a step‑by‑step tutorial for setting a ListBox's selected index, refreshing it, and verifying the selection with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeRefreshDemo
{
    // Demonstrates how to synchronize a ListBox (or other form control) with its linked cell after the cell value changes. The example creates a workbook, adds a ListBox linked to a range, sets the linked cell, calls UpdateSelectedValue on the shape and on the worksheet's Shapes collection, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for a ListBox control
            sheet.Cells["A1"].Value = "Option 1";
            sheet.Cells["A2"].Value = "Option 2";
            sheet.Cells["A3"].Value = "Option 3";

            // Add a ListBox shape to the worksheet
            // Parameters: upperRow, leftColumn, upperOffset, leftOffset, height, width
            Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Set the range that provides the list items
            listBoxShape.SetInputRange("$A$1:$A$3", false, false);

            // Link the selected value of the ListBox to cell B1
            listBoxShape.SetLinkedCell("$B$1", false, true);

            // Initial selection: set linked cell to the second item (value "Option 2")
            sheet.Cells["B1"].Value = "Option 2";

            // Refresh the shape so it reflects the current linked cell value
            listBoxShape.UpdateSelectedValue();

            // Verify the selection (optional)
            ListBox listBox = (ListBox)listBoxShape;
            Console.WriteLine("Initially selected: " + (listBox.IsSelected(1) ? "Option 2" : "None"));

            // Change the linked cell value to "Option 3"
            sheet.Cells["B1"].Value = "Option 3";

            // Refresh all shapes in the worksheet (updates the ListBox selection)
            sheet.Shapes.UpdateSelectedValue();

            // Verify the new selection
            Console.WriteLine("After update selected: " + (listBox.IsSelected(2) ? "Option 3" : "None"));

            // Save the workbook
            workbook.Save("LinkedShapeRefreshDemo.xlsx");
        }
    }
}
