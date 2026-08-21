// Title: Add a ListBox shape linked to a data‑validation list in Aspose.Cells for .NET
// Description: Creates a workbook, fills B1‑B3 with options, adds a list‑type validation to A1, inserts a ListBox shape, sets its input range, links it to A1, updates the displayed selection, and saves the file.
// Keywords: Aspose.Cells ListBox shape | C# data validation linking | shape linked cell Aspose | UpdateSelectedValue method | SetInputRange ListBox | Excel form controls programmatically | .NET workbook shape binding
// Common Searches: Aspose.Cells bind ListBox to validation list | C# sync shape text with cell value | how to link shape to cell in Aspose.Cells | update ListBox selection from cell Aspose | programmatic data validation with shapes .NET
// Developer Intent: Generate a ListBox control that reflects a cell’s validation list and stays synchronized with the linked cell.
// Use Cases: Interactive worksheets where a visual ListBox mirrors a drop‑down cell. | Report templates that display the chosen validation option inside a shape. | Excel‑style forms that automatically update shape text when the underlying cell changes.
// AI Prompts: Write C# code using Aspose.Cells to add a ListBox shape, set its input range, link it to a validation cell, and keep the displayed text in sync. | Explain the role of SetLinkedCell and UpdateSelectedValue when connecting a shape to a data‑validation list in Aspose.Cells. | Provide step‑by‑step instructions for creating a validation list, inserting a ListBox, and binding the shape to the same range with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeValidationDemo
{
    // Creates a workbook, fills B1‑B3 with options, adds a list‑type validation to A1, inserts a ListBox shape, sets its input range, links it to A1, updates the displayed selection, and saves the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Prepare the list of options in cells B1:B3
                // -------------------------------------------------
                worksheet.Cells["B1"].Value = "Option1";
                worksheet.Cells["B2"].Value = "Option2";
                worksheet.Cells["B3"].Value = "Option3";

                // -------------------------------------------------
                // 2. Add data validation to cell A1 that uses the list in B1:B3
                // -------------------------------------------------
                // Add a new validation rule (GetValidation returns null if none exists)
                int validationIndex = worksheet.Validations.Add();
                Validation validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                // Use a range reference for the list items
                validation.Formula1 = "$B$1:$B$3";
                validation.InCellDropDown = true; // Show dropdown in the cell

                // -------------------------------------------------
                // 3. Add a ListBox shape and link it to the same validation list
                // -------------------------------------------------
                // Parameters: upper left row, upper left column, top, left, width, height
                Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

                // Fill the ListBox with the same range used for validation
                listBoxShape.SetInputRange("$B$1:$B$3", false, false);

                // Link the ListBox to cell A1 so its selected value reflects the cell value
                listBoxShape.SetLinkedCell("$A$1", false, true);

                // -------------------------------------------------
                // 4. Synchronize the shape's displayed text with the linked cell
                // -------------------------------------------------
                // This updates the selected item in the ListBox based on the current value of A1
                listBoxShape.UpdateSelectedValue();

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save("ShapeLinkedToValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
