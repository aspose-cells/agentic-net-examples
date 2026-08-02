// Title: C# – Add a Shape, Link It to a Cell Containing a Custom Array Constant, and Display the Values with Aspose.Cells
// Description: Demonstrates how to create a new workbook with Aspose.Cells for .NET, define a string array constant ("Apple","Banana","Cherry") as an array formula in B1, calculate the formulas, add a rectangle shape, link the shape to the first cell of the array using SetLinkedCell, refresh the shape’s displayed value, output the linked cell address and each array element, and save the file as ShapeLinkedArrayDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape linking | SetLinkedCell | array formula | custom array constant | rectangle shape | UpdateSelectedValue | Excel automation | linked cell value
// Common Searches: Aspose.Cells link shape to cell array constant C# | SetLinkedCell with custom array formula Aspose.Cells | Add rectangle shape and bind to cell in .NET | How to display array constant values in a linked shape | Update shape value after changing linked cells Aspose.Cells
// Developer Intent: Create a shape, bind it to a cell that holds a custom array constant, and retrieve or display the linked values programmatically.
// Use Cases: Generate dynamic report headers where a shape shows the first item of a predefined list. | Build interactive dashboards that reflect the current value of a calculated array range via a linked shape. | Automate Excel templates where shapes display category names taken from a constant array.
// AI Prompts: Write C# code using Aspose.Cells to set a string array constant as an array formula and link a rectangle shape to the first cell of that array. | Explain the purpose of the isR1C1 and isLocal parameters in Shape.SetLinkedCell and how they influence the linking behavior. | Provide steps to refresh a shape’s displayed value after modifying the values in the linked array cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook with Aspose.Cells for .NET, define a string array constant ("Apple","Banana","Cherry") as an array formula in B1, calculate the formulas, add a rectangle shape, link the shape to the first cell of the array using SetLinkedCell, refresh the shape’s displayed value, output the linked cell address and each array element, and save the file as ShapeLinkedArrayDemo.xlsx.
class ShapeLinkedArrayDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define a custom array constant {"Apple","Banana","Cherry"}
        // Set it as an array formula starting at B1, spilling to 1 row and 3 columns
        Cell arrayCell = cells["B1"];
        // The formula must start with '=' and use curly braces for the constant
        arrayCell.SetArrayFormula("={\"Apple\",\"Banana\",\"Cherry\"}", 1, 3);

        // Calculate formulas so the array values are materialized in the cells
        workbook.CalculateFormula();

        // Add a rectangle shape (could be any shape) to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 50);

        // Link the shape's value to the first cell of the array (B1)
        // isR1C1 = false (A1 style), isLocal = false (default locale)
        shape.SetLinkedCell("B1", false, false);

        // Optionally update the shape's selected value from the linked cell
        sheet.Shapes.UpdateSelectedValue();

        // Display the linked cell address and the array values in the console
        Console.WriteLine("Shape linked cell: " + shape.LinkedCell);
        Console.WriteLine("Array values populated in the worksheet:");
        for (int col = 0; col < 3; col++)
        {
            Cell c = cells[0, 1 + col]; // B1, C1, D1
            Console.WriteLine($"{c.Name}: {c.StringValue}");
        }

        // Save the workbook to a file
        workbook.Save("ShapeLinkedArrayDemo.xlsx");
    }
}
