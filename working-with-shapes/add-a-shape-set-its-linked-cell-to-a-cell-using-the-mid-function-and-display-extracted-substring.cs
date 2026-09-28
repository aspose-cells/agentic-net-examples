// Title: Create a rectangle shape in an Excel worksheet and bind its text to a MID formula using Aspose.Cells for .NET (C#)
// AI Prompts: Add a rectangle shape to a worksheet and set its Text property to a MID formula that extracts a substring from cell A1 with Aspose.Cells in C#. | Generate an Excel file where a shape displays the result of =MID(A1,2,5) by linking the shape's text to the formula using Aspose.Cells .NET API. | Create C# code that adds a shape, assigns a MID formula to its caption, and saves the workbook.
// Common Searches: Aspose.Cells example showing shape displaying MID(A1,2,5) result | how to bind a rectangle shape to a cell formula in C# | display substring from a cell inside a shape using Aspose.Cells .NET | C# code to create shape with linked formula in Excel workbook
// Tags: add rectangle shape Aspose.Cells | link shape text to cell formula | display MID function result in shape | Aspose.Cells shape binding example | C# generate workbook with linked shape

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Put sample text into cell A1 (the source for the MID function)
sheet.Cells["A1"].PutValue("Aspose.Cells");

// Add a rectangle shape to the worksheet
// Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
Aspose.Cells.Drawing.Shape shape = sheet.Shapes.AddShape(
    Aspose.Cells.Drawing.MsoDrawingType.Rectangle, // shape type
    2,    // upper left row
    0,    // upper left column
    0,    // top offset (in points)
    0,    // left offset (in points)
    100,  // height (in points)
    30    // width (in points)
);

// Link the shape's displayed text to a cell using the MID function
// The formula extracts a substring from A1; the shape will show the result
shape.Text = "=MID(A1,2,5)"; // extracts "spose"

// Save the workbook to a file
workbook.Save("ShapeLinkedCell.xlsx");
