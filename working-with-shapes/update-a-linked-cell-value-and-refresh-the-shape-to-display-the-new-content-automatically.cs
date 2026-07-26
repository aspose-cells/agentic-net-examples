// Title: Refresh a ListBox shape after changing its linked cell using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a ListBox shape, bind it to a range, link it to a cell, modify the linked cell value, and call UpdateSelectedValue so the shape instantly reflects the new selection.
// Keywords: Aspose.Cells ListBox shape | linked cell refresh | UpdateSelectedValue .NET | programmatic shape selection | Excel form control automation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells update ListBox after linked cell change | How to sync ListBox shape with cell value in .NET | SetLinkedCell and refresh shape Aspose.Cells | Refresh Excel form controls programmatically | UpdateSelectedValue usage example
// Developer Intent: Synchronize a ListBox shape with its linked cell so the displayed selection updates automatically after the cell value changes.
// Use Cases: Automatically adjust a form control based on calculation results stored in a worksheet cell. | Create interactive reports where ListBox selections reflect dynamic data without manual refresh. | Implement UI logic that changes ListBox items programmatically during workbook generation.
// AI Prompts: Generate C# code that binds a ListBox shape to a range and updates its selection when the linked cell value changes using Aspose.Cells. | Show an example of refreshing multiple form controls after modifying their linked cells in an Aspose.Cells workbook. | Explain the purpose of UpdateSelectedValue and when it should be invoked after editing a linked cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeDemo
{
    // Demonstrates how to create a workbook, add a ListBox shape, bind it to a range, link it to a cell, modify the linked cell value, and call UpdateSelectedValue so the shape instantly reflects the new selection.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the ListBox input range
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[i, 0].PutValue(i + 1); // A1:A6 = 1..6
            }

            // Add a ListBox shape to the worksheet
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Define the input range (items) and the linked cell (selected value)
            listBoxShape.SetInputRange("$A$1:$A$6", false, false);
            listBoxShape.SetLinkedCell("$A$12", false, true);

            // Set an initial value in the linked cell and refresh the shape
            worksheet.Cells["A12"].PutValue(3);
            listBoxShape.UpdateSelectedValue(); // Refresh selection based on linked cell

            // Change the linked cell value and refresh again
            worksheet.Cells["A12"].PutValue(5);
            listBoxShape.UpdateSelectedValue(); // Shape now reflects the new selection

            // Save the workbook to verify the result
            workbook.Save("LinkedShapeDemo.xlsx");
        }
    }
}
