// Title: Refresh linked ComboBox and other shapes after saving a workbook – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a ComboBox linked to cells A1:A3, modify source data, and call sheet.Shapes.UpdateSelectedValue() to synchronize the shape's selected value before the final save, ensuring accurate linked shape positions and selections.
// Keywords: Aspose.Cells | C# | Refresh linked shapes | UpdateSelectedValue | ComboBox shape | worksheet shapes | Excel automation | save workbook | linked control | shape synchronization
// Common Searches: Aspose.Cells update linked shape after cell change | C# Refresh ComboBox selected value in Excel | UpdateSelectedValue method example | Synchronize linked shapes after saving workbook | How to refresh linked shapes in Aspose.Cells | Refresh shape positions after workbook save
// Developer Intent: Synchronize the selected values of linked shapes (e.g., ComboBox) with their source cells after modifying data and saving the workbook.
// Use Cases: Generate reports where dropdowns must reflect programmatically changed options before distribution. | Maintain UI consistency in Excel templates that use linked shapes after batch data updates. | Ensure multiple linked shapes stay in sync when source ranges are edited during automated processing.
// AI Prompts: Write C# code using Aspose.Cells to refresh all linked shapes after updating their source cells. | Show how to use the UpdateSelectedValue method for ComboBox, ListBox, and CheckBox shapes in a worksheet. | Explain the steps to keep linked shape selections accurate when saving a workbook multiple times with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace RefreshLinkedShapesDemo
{
    // Demonstrates how to create a workbook, add a ComboBox linked to cells A1:A3, modify source data, and call sheet.Shapes.UpdateSelectedValue() to synchronize the shape's selected value before the final save, ensuring accurate linked shape positions and selections.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // Add some data that will be linked to a shape
            sheet.Cells["A1"].PutValue("Option 1");
            sheet.Cells["A2"].PutValue("Option 2");
            sheet.Cells["A3"].PutValue("Option 3");

            // Add a ComboBox shape and link its input range to the data above
            ComboBox combo = (ComboBox)sheet.Shapes.AddComboBox(2, 2, 120, 20, 3, 20);
            combo.SetInputRange("A1:A3", false, false);            // link range
            combo.SelectedIndex = 0;                               // initial selection

            // Save the workbook (first save)
            workbook.Save("LinkedShapesBeforeRefresh.xlsx");        // save

            // Change the linked cell value to see the effect after refresh
            sheet.Cells["A2"].PutValue("Updated Option 2");

            // Save again (second save)
            workbook.Save("LinkedShapesAfterChange.xlsx");          // save

            // Refresh linked shapes so that their selected values reflect the current cell contents
            // This updates all shapes in the worksheet collection
            sheet.Shapes.UpdateSelectedValue();                     // refresh linked shapes

            // Optional: Save the workbook after refreshing to persist the updated shape state
            workbook.Save("LinkedShapesRefreshed.xlsx");            // save
        }
    }
}
