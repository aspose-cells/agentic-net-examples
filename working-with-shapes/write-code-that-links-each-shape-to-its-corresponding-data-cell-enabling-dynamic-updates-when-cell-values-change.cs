// Title: Aspose.Cells C# – Link ListBox, CheckBox, and ScrollBar Shapes to Worksheet Cells for Real‑Time Updates
// Description: Demonstrates how to create a workbook, add ListBox, CheckBox, and ScrollBar shapes, bind each shape to a specific cell using SetLinkedCell (and SetInputRange for the ListBox), refresh the visual state with UpdateSelectedValue, modify the linked cells programmatically, and save the result.
// Keywords: Aspose.Cells shape linking | C# SetLinkedCell example | dynamic shape update Aspose.Cells | ListBox shape bound to cell | CheckBox shape linked cell | ScrollBar shape cell binding | UpdateSelectedValue Aspose.Cells | .NET spreadsheet shape binding
// Common Searches: how to bind a ListBox shape to a cell using Aspose.Cells for .NET | refresh Aspose.Cells shapes after changing linked cell values | set linked cell for CheckBox shape in Aspose.Cells C# | update ScrollBar shape when numeric cell changes | Aspose.Cells shape to cell synchronization
// Developer Intent: Bind each form control shape to a worksheet cell so the shape automatically reflects any changes made to the cell value and vice‑versa.
// Use Cases: Synchronize a ListBox shape with a data range and a selected‑index cell, then programmatically change the index and refresh the shape. | Tie a CheckBox shape to a Boolean cell, allowing user clicks or code to keep the cell and shape in sync. | Connect a ScrollBar shape to a numeric cell, modify the cell value in code, and call UpdateSelectedValue to move the thumb accordingly.
// AI Prompts: Generate C# code that adds a ComboBox shape, sets its input range and linked cell, and updates the shape after the linked cell value changes using Aspose.Cells. | Show how to batch‑update linked cells for multiple shapes (ListBox, CheckBox, ScrollBar) and efficiently refresh all shapes in a worksheet with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // Demonstrates how to create a workbook, add ListBox, CheckBox, and ScrollBar shapes, bind each shape to a specific cell using SetLinkedCell (and SetInputRange for the ListBox), refresh the visual state with UpdateSelectedValue, modify the linked cells programmatically, and save the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the controls
                // Column A: items for the ListBox
                for (int i = 0; i < 5; i++)
                    sheet.Cells[i, 0].Value = $"Item {i + 1}";

                // B1: selected index for the ListBox (zero‑based)
                sheet.Cells["B1"].Value = 2;

                // C1: state for the CheckBox (TRUE/FALSE)
                sheet.Cells["C1"].Value = true;

                // D1: value for the ScrollBar
                sheet.Cells["D1"].Value = 30;

                // ---------- Add shapes ----------
                // 1. ListBox shape
                Shape listBoxShape = sheet.Shapes.AddListBox(2, 2, 100, 100, 5, 20);
                listBoxShape.SetInputRange("$A$1:$A$5", false, false);
                listBoxShape.SetLinkedCell("$B$1", false, true);

                // 2. CheckBox shape (index 0, count 1)
                Shape checkBoxShape = sheet.Shapes.AddCheckBox(2, 5, 100, 20, 0, 1);
                checkBoxShape.SetLinkedCell("$C$1", false, true);

                // 3. ScrollBar shape (index 0, count 1)
                Shape scrollBarShape = sheet.Shapes.AddScrollBar(2, 8, 150, 20, 0, 1);
                scrollBarShape.SetLinkedCell("$D$1", false, true);
                scrollBarShape.SetInputRange("$D$1:$D$1", false, false); // optional illustration

                // Ensure visual state matches linked cells
                sheet.Shapes.UpdateSelectedValue();

                // ---------- Demonstrate dynamic update ----------
                sheet.Cells["B1"].Value = 4;          // select fifth item in ListBox
                sheet.Cells["C1"].Value = false;     // uncheck the CheckBox
                sheet.Cells["D1"].Value = 70;        // move ScrollBar thumb

                // Refresh shapes to reflect new values
                sheet.Shapes.UpdateSelectedValue();

                // Save the workbook
                string outputPath = "ShapesLinkedToCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
