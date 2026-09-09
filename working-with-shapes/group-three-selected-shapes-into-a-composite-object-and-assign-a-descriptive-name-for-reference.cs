// Title: How to group multiple worksheet shapes into a named GroupShape using Aspose.Cells for .NET (C#)
// AI Prompts: Create three rectangle shapes on a worksheet, group them with Worksheet.Shapes.Group, and set the GroupShape.Name to a custom identifier. | Combine selected shapes into a single GroupShape object in C# and retrieve it later by its assigned name. | Generate a composite shape from multiple Excel shapes using Aspose.Cells and assign a descriptive name for future reference.
// Common Searches: Aspose.Cells C# group several shapes into one GroupShape and name it | How to assign a custom name to a grouped shape in an Excel workbook using Aspose.Cells | C# code to create a composite shape from multiple rectangles in Aspose.Cells
// Tags: group shapes Aspose.Cells C# | GroupShape naming Aspose.Cells | composite shape creation Excel .NET | worksheet shape grouping example | assign custom name to GroupShape

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();               // create workbook
Worksheet worksheet = workbook.Worksheets[0];    // get first sheet

// Add three sample shapes (rectangles) to the worksheet
Shape shape1 = worksheet.Shapes.AddShape(
    MsoDrawingType.Rectangle, // shape type
    1,   // upper left row
    1,   // upper left column
    0,   // top offset in pixels
    0,   // left offset in pixels
    100, // width in pixels
    50); // height in pixels

Shape shape2 = worksheet.Shapes.AddShape(
    MsoDrawingType.Rectangle,
    2,
    1,
    0,
    0,
    100,
    50);

Shape shape3 = worksheet.Shapes.AddShape(
    MsoDrawingType.Rectangle,
    3,
    1,
    0,
    0,
    100,
    50);

// Group the three shapes into a composite (GroupShape) object
Shape[] shapesToGroup = new Shape[] { shape1, shape2, shape3 };
GroupShape compositeShape = worksheet.Shapes.Group(shapesToGroup);

// Assign a descriptive name to the composite shape for later reference
compositeShape.Name = "MyCompositeShape";

// Save the workbook to a file
workbook.Save("GroupedShapes.xlsx");
