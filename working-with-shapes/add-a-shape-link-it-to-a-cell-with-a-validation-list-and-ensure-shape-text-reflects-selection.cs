// Title: Add a ListBox shape linked to a data‑validation cell and sync its text – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate a range with options, add an in‑cell dropdown validation to B1, insert a ListBox shape, set the shape's input range, link the shape to the validation cell, initialize the cell value, update the shape's selected item, verify the selection, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# ListBox shape | link shape to validation cell | Excel shape linked cell | update ListBox selected value | set input range Aspose.Cells | data validation dropdown Aspose | sync shape text with cell | Aspose.Cells .NET example
// Common Searches: Aspose.Cells how to bind ListBox shape to a cell | C# link Excel shape to data validation list | Aspose.Cells update ListBox after cell change | Create ListBox shape with linked cell in .NET | Synchronize shape text with validation cell Aspose
// Developer Intent: Create a ListBox shape, link it to a cell that contains a data‑validation list, and keep the shape’s displayed text synchronized with the cell’s selection.
// Use Cases: Provide a visual dropdown outside the cell while using the same validation source. | Programmatically set the initial ListBox selection based on a cell value and reflect later changes automatically. | Validate UI consistency by checking that the shape’s selected index matches the linked cell.
// AI Prompts: Generate C# code with Aspose.Cells to add a ListBox shape linked to a validation cell and synchronize its selected value. | Explain the role of SetLinkedCell and UpdateSelectedValue when syncing a shape with a cell in Aspose.Cells. | Show error‑handling patterns for creating and linking shapes in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, populate a range with options, add an in‑cell dropdown validation to B1, insert a ListBox shape, set the shape's input range, link the shape to the validation cell, initialize the cell value, update the shape's selected item, verify the selection, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate a range with list items for validation and the shape
            worksheet.Cells["A1"].Value = "Option1";
            worksheet.Cells["A2"].Value = "Option2";
            worksheet.Cells["A3"].Value = "Option3";

            // Create a data validation on cell B1 with an in‑cell dropdown list
            Validation validation = worksheet.Cells["B1"].GetValidation();
            validation.Type = ValidationType.List;
            validation.Formula1 = "$A$1:$A$3";
            validation.InCellDropDown = true;

            // Add a ListBox shape to the worksheet
            Shape listBoxShape = worksheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);
            if (listBoxShape == null)
                throw new InvalidOperationException("Failed to create ListBox shape.");

            // Set the shape's input range to the same list of options
            listBoxShape.SetInputRange("$A$1:$A$3", false, false);

            // Link the shape to the validation cell (B1) so its value reflects the cell's content
            listBoxShape.SetLinkedCell("$B$1", false, true);

            // Set an initial value in the linked cell; the shape will display this option
            worksheet.Cells["B1"].Value = "Option2";

            // Update the shape's selected value based on the linked cell
            listBoxShape.UpdateSelectedValue();

            // Verify the selection programmatically
            if (listBoxShape is ListBox listBox && listBox.IsSelected(1)) // zero‑based index; 1 corresponds to "Option2"
            {
                Console.WriteLine("Shape correctly shows Option2.");
            }

            // Define output file path
            string outputPath = "ShapeLinkedValidationDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
