// Title: Link a Rectangle Shape to a Cell with Custom Date Formatting and Validate the Date using Aspose.Cells for .NET
// Description: Demonstrates how to insert the current date into cell B2, apply the custom format "dd-mmm-yyyy", add a rectangle shape, link the shape to the formatted cell, refresh the displayed value, verify that the linked cell holds a DateTime object, and save the workbook as ShapeLinkedDate.xlsx.
// Keywords: Aspose.Cells | .NET | shape linked cell | custom date format | DateTime validation | rectangle shape | Excel shape linking | SetLinkedCell | UpdateSelectedValue | Aspose.Cells example
// Common Searches: Aspose.Cells link shape to cell with date | apply custom date format to Excel cell using Aspose.Cells | verify DateTime value of linked cell in Aspose.Cells | add rectangle shape and bind it to a cell in .NET | refresh shape text after linking to a cell
// Developer Intent: Create a rectangle shape that displays a custom‑formatted date from a linked cell and confirm the cell contains a valid DateTime value.
// Use Cases: Generate a report where a shape shows the generation date in a specific format. | Build an interactive dashboard with shapes that automatically reflect date changes in linked cells. | Validate imported worksheet data by checking that linked cells hold proper DateTime types before further processing.
// AI Prompts: Write C# code with Aspose.Cells to add a rectangle shape, link it to cell B2, format the cell as dd-mmm-yyyy, update the shape's text, and confirm the linked value is a DateTime. | Show an example that inserts the current date into a cell, applies a custom date format, links a shape to that cell, refreshes the shape, and saves the workbook. | Explain how to use SetLinkedCell and UpdateSelectedValue to bind a shape to a date cell and how to verify the cell's data type in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to insert the current date into cell B2, apply the custom format "dd-mmm-yyyy", add a rectangle shape, link the shape to the formatted cell, refresh the displayed value, verify that the linked cell holds a DateTime object, and save the workbook as ShapeLinkedDate.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put the current date into cell B2
            Cell dateCell = sheet.Cells["B2"];
            dateCell.PutValue(DateTime.Now);

            // Apply custom date format "dd-mmm-yyyy" to the cell
            Style style = workbook.CreateStyle();
            style.Custom = "dd-mmm-yyyy";
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // enable number format changes
            dateCell.SetStyle(style, flag);

            // Add a rectangle shape (acts as a text box)
            // Parameters: upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, width, height
            RectangleShape shape = sheet.Shapes.AddRectangle(5, 2, 8, 6, 100, 50);
            shape.Text = "Date:"; // initial placeholder text

            // Link the shape to cell B2 and refresh displayed value
            shape.SetLinkedCell("$B$2", false, true);
            shape.UpdateSelectedValue();

            // Verify that the linked cell contains a DateTime value
            object linkedValue = dateCell.Value;
            if (linkedValue is DateTime dt)
            {
                Console.WriteLine("Linked cell contains a valid date: " + dt.ToString("dd-MMM-yyyy"));
            }
            else
            {
                Console.WriteLine("Linked cell does not contain a date.");
            }

            // Save the workbook
            string outputPath = "ShapeLinkedDate.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
