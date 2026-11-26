// Title: Clone a TextBox shape, change its placeholder text, and insert it into a new row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells clone textbox shape and set new text | How to copy a TextBox to another row using Aspose.Cells Shapes.AddCopy | Aspose.Cells example for duplicating a textbox and updating placeholder | Insert cloned TextBox into specific row in Excel with Aspose.Cells .NET | Change placeholder text of a copied TextBox shape in Aspose.Cells workbook | Aspose.Cells C# addcopy textbox to different row example
// Tags: textbox shape cloning Aspose.Cells | Shapes.AddCopy method C# | update textbox placeholder Aspose.Cells | clone textbox to new row Excel .NET | modify cloned textbox visual properties Aspose.Cells | duplicate shape with formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextBoxCloneDemo
{
    // This C# example creates a template TextBox in a worksheet, clones it to a different row using the Shapes.AddCopy method, updates the cloned TextBox's placeholder text and visual styling, and saves the workbook as an Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Create a template TextBox (the source that will be cloned)
                // ------------------------------------------------------------
                // Add a TextBox at row 1, column 1 with width=200 pixels and height=80 pixels
                int templateIndex = worksheet.TextBoxes.Add(1, 1, 200, 80);
                TextBox templateBox = worksheet.TextBoxes[templateIndex];

                // Set some initial properties for the template
                templateBox.Text = "Template Placeholder";
                templateBox.Font.Size = 12;
                templateBox.Font.Color = System.Drawing.Color.DarkBlue;
                templateBox.Fill.SolidFill.Color = System.Drawing.Color.LightYellow;
                templateBox.Line.Weight = 1.5;
                templateBox.Line.DashStyle = MsoLineDashStyle.Solid;
                templateBox.Placement = PlacementType.MoveAndSize;

                // ------------------------------------------------------------
                // 2. Clone the template TextBox
                // ------------------------------------------------------------
                // Determine the size of the template (in rows/columns)
                int templateHeightRows = templateBox.LowerRightRow - templateBox.UpperLeftRow;
                int templateWidthCols = templateBox.LowerRightColumn - templateBox.UpperLeftColumn;

                // Define the target location (new row) for the cloned TextBox
                int targetRow = 10;          // place the clone starting at row 10
                int targetColumn = templateBox.UpperLeftColumn; // keep the same column as the template

                // Use the Shapes collection to clone the shape (TextBox inherits from Shape)
                Shape clonedShape = worksheet.Shapes.AddCopy(
                    templateBox,                     // source shape to copy
                    targetRow,                       // upper left row of the copy
                    targetColumn,                    // upper left column of the copy
                    targetRow + templateHeightRows,   // lower right row
                    targetColumn + templateWidthCols // lower right column
                );

                // Cast the cloned shape back to TextBox
                TextBox clonedBox = (TextBox)clonedShape;

                // ------------------------------------------------------------
                // 3. Update the placeholder text of the cloned TextBox
                // ------------------------------------------------------------
                clonedBox.Text = "Cloned Placeholder";

                // Optionally adjust other visual properties of the clone
                clonedBox.Font.Color = System.Drawing.Color.Maroon;
                clonedBox.Fill.SolidFill.Color = System.Drawing.Color.LightCyan;

                // ------------------------------------------------------------
                // 4. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "ClonedTextBoxDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
