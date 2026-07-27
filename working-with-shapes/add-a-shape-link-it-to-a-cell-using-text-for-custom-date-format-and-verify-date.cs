// Title: C# – Add a Rectangle Shape Linked to a Date Cell with Custom Format and Verify It using Aspose.Cells
// Description: Demonstrates how to create a workbook, write a DateTime to cell B2, apply a custom "dd-MMM-yyyy" format, insert a rectangle shape, link the shape to the cell without absolute references, update and read the displayed value, compare it to the expected format, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | link shape to cell | custom date format | rectangle shape | SetLinkedCell | GetLinkedCell | shape verification | Excel automation
// Common Searches: Aspose.Cells link shape to cell C# | how to display formatted date in a shape using Aspose.Cells | verify shape text matches linked cell value .NET | SetLinkedCell without $ signs Aspose.Cells | add rectangle shape and bind to date cell Aspose.Cells
// Developer Intent: Create a rectangle shape, bind it to a date cell with a custom format, and programmatically confirm that the shape shows the correctly formatted date.
// Use Cases: Show an invoice or due date inside a movable shape on a report sheet. | Build a dashboard where shapes automatically reflect dates from cells using a specific display format. | Validate that shape captions match cell formatting before publishing the workbook.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape, link it to a date cell using a custom format, and verify the displayed text. | Explain the role of the absolute/relative flags in SetLinkedCell when linking a shape to a cell. | Provide a reusable method that compares a shape's displayed value with the formatted string of its linked cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, write a DateTime to cell B2, apply a custom "dd-MMM-yyyy" format, insert a rectangle shape, link the shape to the cell without absolute references, update and read the displayed value, compare it to the expected format, and save the file with Aspose.Cells for .NET.
class ShapeLinkDateExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a date value in cell B2
            Cell dateCell = sheet.Cells["B2"];
            dateCell.PutValue(new DateTime(2023, 12, 25));

            // Apply a custom date format (e.g., "dd-mmm-yyyy")
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";
            dateCell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 1, 0, 0, 100, 50);
            rect.Name = "DateShape";

            // Link the shape to the cell B2 using A1 style formula (without $ signs)
            rect.SetLinkedCell("B2", false, true);

            // Update the shape's displayed value based on the linked cell
            rect.UpdateSelectedValue();

            // Retrieve the linked cell address (same absolute/relative settings as above)
            string linkedCellAddress = rect.GetLinkedCell(false, true);
            Cell linkedCell = sheet.Cells[linkedCellAddress];
            string displayedValue = linkedCell.StringValue; // This respects the custom format

            Console.WriteLine($"Shape is linked to cell: {linkedCellAddress}");
            Console.WriteLine($"Cell value (formatted): {displayedValue}");

            // Simple verification: check if the formatted string matches the expected format
            string expected = dateCell.DateTimeValue.ToString("dd-MMM-yyyy");
            if (displayedValue.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Verification succeeded: date format matches.");
            }
            else
            {
                Console.WriteLine($"Verification failed: expected '{expected}' but got '{displayedValue}'.");
            }

            // Save the workbook
            workbook.Save("ShapeLinkedDate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
