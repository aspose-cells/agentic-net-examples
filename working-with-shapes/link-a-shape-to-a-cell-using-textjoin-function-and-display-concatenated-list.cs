// Title: How to bind a rectangle shape to a TEXTJOIN result cell in Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, fill column A with values, set a TEXTJOIN formula in B1, calculate the workbook, add a rectangle shape, and assign the shape's text to the evaluated TEXTJOIN result using Aspose.Cells C#. | Use Aspose.Cells to insert a rectangle shape at a specific location and programmatically link its displayed text to the concatenated string produced by a TEXTJOIN formula in another cell.
// Common Searches: Aspose.Cells C# set shape text from cell formula result | link rectangle shape to TEXTJOIN output in Excel using Aspose.Cells | display concatenated list from column A in a shape with Aspose.Cells .NET | how to bind shape text to calculated cell value Aspose.Cells | Aspose.Cells example TEXTJOIN and shape text binding
// Tags: Aspose.Cells rectangle shape text binding | C# TEXTJOIN formula example | link shape to calculated cell value Aspose.Cells | set shape text from workbook formula | display concatenated list in Excel shape .NET

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Fill some sample data in column A (A1:A5)
sheet.Cells["A1"].PutValue("Apple");
sheet.Cells["A2"].PutValue("Banana");
sheet.Cells["A3"].PutValue("Cherry");
sheet.Cells["A4"].PutValue("Date");
sheet.Cells["A5"].PutValue("Elderberry");

// Set a TEXTJOIN formula in cell B1 to concatenate the values from A1:A5
sheet.Cells["B1"].Formula = "TEXTJOIN(\", \", TRUE, A1:A5)";

// Calculate formulas so that B1 contains the concatenated result
workbook.CalculateFormula();

// Insert a rectangle shape that will display the concatenated list
// Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
Aspose.Cells.Drawing.Shape shape = sheet.Shapes.AddShape(
    Aspose.Cells.Drawing.MsoDrawingType.Rectangle, // shape type
    2,    // upper left row (row index starts at 0, so row 2 = third row)
    0,    // upper left column
    5,    // top offset (points)
    5,    // left offset (points)
    200,  // height (points)
    300   // width (points)
);

// Set the shape's text to the value calculated in B1
shape.Text = sheet.Cells["B1"].StringValue;

// Optionally, format the shape's text (e.g., font size)
shape.Font.Size = 12;

// Save the workbook to a file
workbook.Save("LinkedShapeWithTextJoin.xlsx");
