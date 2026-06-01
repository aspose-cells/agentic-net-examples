using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();               // create
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Sample shapes – in real scenario shapes already exist
        // -------------------------------------------------
        // Add a rectangle shape
        Shape rectShape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 120, 60);
        // Add a non‑rectangle shape (e.g., oval) for contrast
        Shape ovalShape = sheet.Shapes.AddOval(2, 0, 2, 0, 80, 80);

        // -------------------------------------------------
        // Filter shapes and apply custom margins to rectangles only
        // -------------------------------------------------
        foreach (Shape shape in sheet.Shapes)
        {
            // Identify rectangle shapes by their AutoShapeType
            if (shape.AutoShapeType == AutoShapeType.Rectangle)
            {
                // Access the text alignment object of the shape
                ShapeTextAlignment alignment = shape.TextBody.TextAlignment;

                // Turn off automatic margin calculation
                alignment.IsAutoMargin = false;

                // Set custom margins (values are in points)
                alignment.TopMarginPt = 5.0;      // top margin
                alignment.BottomMarginPt = 5.0;   // bottom margin
                alignment.LeftMarginPt = 8.0;     // left margin
                alignment.RightMarginPt = 8.0;    // right margin
            }
        }

        // Save the modified workbook
        workbook.Save("FilteredRectangles.xlsx");          // save
    }
}