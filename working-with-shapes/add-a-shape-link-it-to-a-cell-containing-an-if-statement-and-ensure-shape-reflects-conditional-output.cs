// Title: C# – Link a Rectangle Shape to an IF Formula Cell with Aspose.Cells
// Description: Demonstrates how to create a workbook, set an IF formula in a cell, add a rectangle shape, link the shape to the formula cell, force the shape to display the calculated result, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | rectangle shape | linked cell | IF formula | dynamic shape text | update shape value | Excel automation | conditional display | shape to cell binding
// Common Searches: Aspose.Cells link shape to cell C# | display IF formula result in shape Aspose.Cells | update shape text after formula recalculation .NET | bind rectangle to worksheet cell Aspose | dynamic shape caption based on Excel formula
// Developer Intent: Add a shape that automatically shows the result of an IF formula by linking it to the formula cell and updating the displayed text.
// Use Cases: Dashboard status indicator that changes between "High" and "Low". | Printable report where shapes act as pass/fail badges driven by formulas. | Interactive worksheet where shape captions reflect real‑time calculations.
// AI Prompts: Generate C# code to change a linked shape's fill color based on the IF result in Aspose.Cells. | Show how to link multiple shapes to different formula cells and synchronize their texts. | Explain how to refresh linked shapes after bulk data updates in an Aspose.Cells workbook.

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Put a value in B1 that will be used by the IF formula
worksheet.Cells["B1"].PutValue(15);

// Set an IF formula in A1 that returns "High" if B1 > 10, otherwise "Low"
worksheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

// Add a rectangle shape to the worksheet (type, upper left row, upper left column, top, left, height, width)
var shape = worksheet.Shapes.AddShape(
    Aspose.Cells.Drawing.MsoDrawingType.Rectangle, // shape type
    2,    // upper left row
    1,    // upper left column
    0,    // top offset (in pixels)
    0,    // left offset (in pixels)
    100,  // height (in points)
    200   // width (in points)
);

// Link the shape to cell A1 so its displayed text reflects the cell's value
shape.SetLinkedCell("A1", true, true);

// Update the shape to reflect the current value of the linked cell
shape.UpdateSelectedValue();

// Save the workbook to a file
workbook.Save("ShapeLinkedToIfFormula.xlsx");
