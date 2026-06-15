using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 50);
            shape.Text = "Sample Text";

            // Configure the shape's text alignment
            var alignment = shape.TextBody.TextAlignment;
            alignment.AutoSize = false;
            alignment.IsTextWrapped = true;
            alignment.RotateTextWithShape = true;
            alignment.TextVerticalOverflow = TextOverflowType.Clip;
            alignment.TextHorizontalOverflow = TextOverflowType.Clip;
            alignment.RotationAngle = 90;
            // Removed TextVerticalType assignment (enum not available in current version)
            alignment.IsLockedText = false;
            alignment.TextShapeType = AutoShapeType.TextBox;
            alignment.TopMarginPt = 2.0;
            alignment.BottomMarginPt = 2.0;
            alignment.LeftMarginPt = 2.0;
            alignment.RightMarginPt = 2.0;
            alignment.IsAutoMargin = true;
            alignment.NumberOfColumns = 1;

            // Save the workbook
            string outputPath = "AlignedShape.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File not found: {fnfEx.FileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}