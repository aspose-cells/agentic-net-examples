// Title: C# Example: Add a Rectangle Shape, Link to a Cell with Value Conversion, and Verify Numeric Result using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert the string "456" into cell B2 as a numeric value, add a rectangle shape, link the shape to B2 with SetLinkedCell, refresh the shape's displayed value, retrieve the linked cell address and numeric value, and save the file as ShapeLinkedCellDemo.xlsx.
// Keywords: Aspose.Cells C# shape linking | SetLinkedCell example | PutValue convert text to number | rectangle shape linked cell | verify linked cell address | numeric value from linked shape | Aspose.Cells workbook demo | C# spreadsheet shape API
// Common Searches: Aspose.Cells link shape to cell C# | convert cell text to number Aspose.Cells | how to use SetLinkedCell with rectangle shape | retrieve linked cell address Aspose.Cells | verify numeric value after linking shape
// Developer Intent: Create a rectangle shape, bind it to a cell whose text is converted to a number, and confirm both the link and the numeric value programmatically.
// Use Cases: Building interactive dashboards where shapes reflect calculated cell values. | Automated testing of spreadsheet templates that require shape‑cell synchronization. | Generating reports that need visual markers linked to numeric data for dynamic updates.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape, links it to cell B2 using SetLinkedCell after converting the cell text to a number, and prints the linked address and numeric value. | Explain the role of each parameter in PutValue and SetLinkedCell when linking a shape to a numeric cell in Aspose.Cells. | Create a unit test in C# that validates a shape's linked cell address and the numeric value stored in the target cell after linking with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert the string "456" into cell B2 as a numeric value, add a rectangle shape, link the shape to B2 with SetLinkedCell, refresh the shape's displayed value, retrieve the linked cell address and numeric value, and save the file as ShapeLinkedCellDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a string that will be converted to a numeric value in cell B2
        // isConverted = true (convert to number), setStyle = false (keep existing style)
        cells["B2"].PutValue("456", true, false);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, width, height, upper left offset X, offset Y
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);

        // Link the shape's value to cell B2
        // formula = "$B$2", isR1C1 = false (A1 style), isLocal = true (locale aware)
        shape.SetLinkedCell("$B$2", false, true);

        // Update the shape's selected value from the linked cell
        shape.UpdateSelectedValue();

        // Verify the linked cell address
        string linkedCellAddress = shape.GetLinkedCell(false, true);
        Console.WriteLine("Shape is linked to cell: " + linkedCellAddress);

        // Verify the numeric value stored in the linked cell
        double numericValue = cells["B2"].DoubleValue;
        Console.WriteLine("Numeric value in linked cell B2: " + numericValue);

        // Save the workbook
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}
