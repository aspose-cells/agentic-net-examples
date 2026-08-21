// Title: Refresh an Aspose.Cells ListBox shape after modifying its linked cell (C#)
// Description: Demonstrates how to create a workbook, add a ListBox shape, assign an input range and a linked cell, set an initial selection, change the linked cell value, and call UpdateSelectedValue so the shape instantly reflects the new selection before saving the file.
// Keywords: Aspose.Cells | C# | ListBox shape | linked cell | UpdateSelectedValue | SetInputRange | SetLinkedCell | shape refresh | Excel form control automation | programmatic selection
// Common Searches: Aspose.Cells refresh ListBox after linked cell change | C# update ListBox linked cell and redraw shape | SetLinkedCell and UpdateSelectedValue example Aspose.Cells | How to sync ListBox shape with cell value in Aspose.Cells
// Developer Intent: Make a ListBox shape display the value stored in its linked cell without reopening the workbook.
// Use Cases: Programmatically set the selected item of a ListBox by writing to the linked cell. | Keep a ListBox in sync with formulas or code that modify the linked cell during workbook generation. | Prepare an Excel file where the ListBox shows the correct selection when the user first opens it.
// AI Prompts: Show C# code using Aspose.Cells to add a ListBox, link it to a cell, change the cell value, and refresh the shape selection. | Provide an example that updates a linked cell and automatically refreshes the associated ListBox shape in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeDemo
{
    // Demonstrates how to create a workbook, add a ListBox shape, assign an input range and a linked cell, set an initial selection, change the linked cell value, and call UpdateSelectedValue so the shape instantly reflects the new selection before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the ListBox input range (A1:A6)
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[i, 0].PutValue(i + 1); // Values 1..6 in column A
            }

            // Add a ListBox shape at row 2, column 0 with size 130x130 pixels
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Define the input range (items) and the linked cell (where the selected index is stored)
            listBoxShape.SetInputRange("$A$1:$A$6", false, false);
            listBoxShape.SetLinkedCell("$A$12", false, true);

            // Set an initial value in the linked cell (e.g., select the third item)
            worksheet.Cells["A12"].PutValue(3);
            // Refresh the shape so it reflects the linked cell value
            listBoxShape.UpdateSelectedValue();

            // Change the linked cell value to select a different item
            worksheet.Cells["A12"].PutValue(5);
            // Refresh again to update the shape's selection
            listBoxShape.UpdateSelectedValue();

            // Save the workbook (the ListBox will show the selected item based on A12)
            workbook.Save("LinkedListBoxDemo.xlsx");
        }
    }
}
