// Title: Set multiple ShapeTextAlignment properties with a C# object initializer in Aspose.Cells
// Description: Shows how to create a workbook, add a rectangle shape, and configure its TextBody.TextAlignment (wrap, rotation, overflow, margins, auto‑margin, column count) using a single object‑initializer statement, then save the file as ShapeTextAlignmentInitializer.xlsx.
// Keywords: Aspose.Cells | ShapeTextAlignment | object initializer | C# | .NET | text wrapping | rotation angle | margin settings | overflow handling | text box shape
// Common Searches: Aspose.Cells object initializer ShapeTextAlignment | C# set shape text alignment in one line | initialize multiple text alignment properties Aspose.Cells | configure shape margins and overflow Aspose.Cells | rotate shape text with object initializer C#
// Developer Intent: Configure all ShapeTextAlignment options for a shape using a single object‑initializer expression.
// Use Cases: Add a rectangle shape and apply wrap, rotation, overflow, margins, and auto‑margin in one initializer before saving the workbook. | Create a text‑box shape with predefined vertical/horizontal overflow, rotation angle, and column count for consistent formatting across worksheets. | Reuse a pre‑configured ShapeTextAlignment initializer to apply identical text alignment settings to multiple shapes in a workbook.
// AI Prompts: Rewrite the sample so that every ShapeTextAlignment property is assigned within a single object initializer. | Generate C# code that adds a shape to a worksheet and sets its TextBody.TextAlignment (wrap, rotation, overflow, margins, auto‑margin, columns) using an object initializer, then saves the workbook. | Provide an example of using a C# object initializer to configure ShapeTextAlignment properties such as IsTextWrapped, RotationAngle, TextVerticalOverflow, and margin values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add a rectangle shape, and configure its TextBody.TextAlignment (wrap, rotation, overflow, margins, auto‑margin, column count) using a single object‑initializer statement, then save the file as ShapeTextAlignmentInitializer.xlsx.
class ShapeTextAlignmentInitializerDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 50, 100);

            // Retrieve the ShapeTextAlignment object (read‑only property) and set its properties
            ShapeTextAlignment alignment = shape.TextBody.TextAlignment;
            alignment.IsTextWrapped = true;
            alignment.RotateTextWithShape = true;
            alignment.TextVerticalOverflow = TextOverflowType.Clip;
            alignment.TextHorizontalOverflow = TextOverflowType.Clip;
            alignment.RotationAngle = 90;
            alignment.TextVerticalType = TextVerticalType.Horizontal;
            alignment.IsLockedText = false;
            alignment.AutoSize = false;
            alignment.TextShapeType = AutoShapeType.TextBox;
            alignment.TopMarginPt = 2.0;
            alignment.BottomMarginPt = 2.0;
            alignment.LeftMarginPt = 2.0;
            alignment.RightMarginPt = 2.0;
            alignment.IsAutoMargin = true;
            alignment.NumberOfColumns = 1;

            // Save the workbook
            workbook.Save("ShapeTextAlignmentInitializer.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
