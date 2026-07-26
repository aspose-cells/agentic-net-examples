// Title: Refresh linked shapes after saving a workbook – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a ComboBox shape linked to a cell range, save the file, and then call `sheet.Shapes.UpdateSelectedValue()` to synchronize the shape's selected value and position. An optional second save persists the refreshed state.
// Keywords: Aspose.Cells | C# | .NET | Refresh linked shapes | UpdateSelectedValue | ComboBox shape | shape synchronization | Workbook.Save | linked shape positions | Aspose.Cells drawing API
// Common Searches: Aspose.Cells refresh linked shapes after save | UpdateSelectedValue C# example | ComboBox shape linked to range Aspose.Cells | Synchronize shape positions after Workbook.Save | How to refresh linked shapes in Aspose.Cells .NET
// Developer Intent: Synchronize the values and positions of cell‑linked shapes after a workbook is saved.
// Use Cases: Ensure a ComboBox reflects the latest range data before persisting the workbook. | Refresh data‑validation drop‑downs or other linked shapes after programmatic changes. | Persist the updated state of linked shapes by saving the workbook a second time.
// AI Prompts: Show a C# code snippet that refreshes all linked shapes after saving a workbook with Aspose.Cells. | Explain why `UpdateSelectedValue` should be called after `Workbook.Save` and how it affects shape positioning. | Provide step‑by‑step instructions to synchronize ComboBox shapes linked to cell ranges in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a ComboBox shape linked to a cell range, save the file, and then call `sheet.Shapes.UpdateSelectedValue()` to synchronize the shape's selected value and position. An optional second save persists the refreshed state.
class RefreshLinkedShapesDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data that will be linked to a shape
        sheet.Cells["B2"].PutValue("Option 1");
        sheet.Cells["B3"].PutValue("Option 2");
        sheet.Cells["B4"].PutValue("Option 3");

        // Add a ComboBox shape and link its input range to the data above
        // (All shape-related operations are part of the drawing API)
        ComboBox combo = (ComboBox)sheet.Shapes.AddComboBox(2, 2, 100, 20, 3, 20);
        combo.SetInputRange("B2:B4", false, false); // link to the range
        combo.SelectedIndex = 0; // initial selection

        // Save the workbook (lifecycle rule: save)
        string filePath = "LinkedShapesDemo.xlsx";
        workbook.Save(filePath);

        // After saving, refresh linked shapes to ensure their positions/values are up‑to‑date
        // The ShapeCollection.UpdateSelectedValue method updates the selected value of all
        // shapes that are linked to cells or ranges.
        sheet.Shapes.UpdateSelectedValue();

        // Optionally, save again if you want the refreshed state persisted
        workbook.Save("LinkedShapesDemo_Refreshed.xlsx");
    }
}
