// Title: Link ListBox, CheckBox, and Spinner Shapes to Worksheet Cells – Aspose.Cells for .NET
// Description: C# sample that creates a workbook, fills columns A‑C, adds ListBox, CheckBox, and Spinner shapes, links each shape to a specific cell, calls UpdateSelectedValue to keep shapes in sync, modifies the cells, refreshes the shapes, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | shape linking | ListBox shape | CheckBox shape | Spinner shape | linked cell | UpdateSelectedValue | dynamic shape update | Excel shape binding | worksheet shapes | cell to shape synchronization
// Common Searches: Aspose.Cells link shape to cell C# | Update shape after changing linked cell Aspose.Cells | Bind ListBox shape to Excel range using Aspose | CheckBox shape linked cell example .NET | Spinner shape cell binding Aspose.Cells | Refresh shapes with UpdateSelectedValue method
// Developer Intent: The developer needs each form control shape to be bound to a worksheet cell so the shape automatically reflects any cell value changes.
// Use Cases: Connect a ListBox shape to a data range and a linked cell, enabling real‑time selection updates in a summary cell. | Bind a CheckBox shape to a Boolean cell to toggle calculations, formatting, or conditional logic based on true/false values. | Link a Spinner shape to a numeric cell, allowing users to increment or decrement a parameter and instantly update dependent formulas.
// AI Prompts: Generate C# code that adds a ComboBox shape, sets its input range, links it to a cell, and synchronizes the selected value after the cell changes using Aspose.Cells. | Show how to loop through rows and create ListBox, CheckBox, and Spinner shapes, each linked to its own cell, with Aspose.Cells for .NET. | Explain the purpose of the UpdateSelectedValue method and best practices for calling it after modifying linked cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    // C# sample that creates a workbook, fills columns A‑C, adds ListBox, CheckBox, and Spinner shapes, links each shape to a specific cell, calls UpdateSelectedValue to keep shapes in sync, modifies the cells, refreshes the shapes, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that will be linked to shapes
            // Column A will hold values for a ListBox, Column B for a CheckBox, Column C for a Spinner
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].Value = i + 1;          // A1:A5
                sheet.Cells[i, 1].Value = (i % 2 == 0);  // B1:B5 (true/false)
                sheet.Cells[i, 2].Value = i * 10;        // C1:C5
            }

            // Add a ListBox shape and link it to cell A10
            Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 120, 120);
            listBoxShape.SetInputRange("$A$1:$A$5", false, false);
            listBoxShape.SetLinkedCell("$A$10", false, true);
            // Initialize linked cell value
            sheet.Cells["A10"].Value = 3;

            // Add a CheckBox shape and link it to cell B10
            Shape checkBoxShape = sheet.Shapes.AddCheckBox(4, 0, 4, 0, 100, 30);
            checkBoxShape.SetLinkedCell("$B$10", false, true);
            sheet.Cells["B10"].Value = true;

            // Add a Spinner shape and link it to cell C10
            Shape spinnerShape = sheet.Shapes.AddSpinner(6, 0, 6, 0, 100, 30);
            spinnerShape.SetLinkedCell("$C$10", false, true);
            sheet.Cells["C10"].Value = 20;

            // Update all shapes so that their selected values reflect the linked cells
            sheet.Shapes.UpdateSelectedValue();

            // Demonstrate dynamic update: change linked cell values and refresh shapes
            sheet.Cells["A10"].Value = 5;   // ListBox should select the 5th item
            sheet.Cells["B10"].Value = false; // CheckBox should become unchecked
            sheet.Cells["C10"].Value = 40; // Spinner should reflect new value

            // Apply the changes to the shapes
            sheet.Shapes.UpdateSelectedValue();

            // Save the workbook
            workbook.Save("ShapesLinkedToCells.xlsx");
        }
    }
}
