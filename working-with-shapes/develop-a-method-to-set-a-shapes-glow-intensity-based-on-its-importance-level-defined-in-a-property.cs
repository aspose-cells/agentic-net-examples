using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeGlowHelper
{
    // Sets the glow effect of a shape based on an importance level (1‑5)
    public static void SetGlowByImportance(Shape shape, int importance)
    {
        // Ensure the importance value is within the expected range
        if (importance < 1) importance = 1;
        if (importance > 5) importance = 5;

        // Map importance to glow size (points). Higher importance → larger glow.
        double[] sizeMap = { 0, 5, 10, 15, 20, 25 }; // index matches importance
        shape.Glow.Size = sizeMap[importance];

        // Choose a glow color that reflects the importance level
        // (lighter colors for low importance, vivid colors for high importance)
        switch (importance)
        {
            case 1:
                shape.Glow.Color.Color = Color.LightGray;
                break;
            case 2:
                shape.Glow.Color.Color = Color.Gray;
                break;
            case 3:
                shape.Glow.Color.Color = Color.DarkGray;
                break;
            case 4:
                shape.Glow.Color.Color = Color.Orange;
                break;
            case 5:
                shape.Glow.Color.Color = Color.Red;
                break;
        }

        // Optional: set a moderate transparency so the glow is visible but not opaque
        shape.Glow.Transparency = 0.4; // 40% transparent
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Example importance level (could be read from a custom property)
        int importanceLevel = 4;

        // Apply the glow effect based on the importance level
        ShapeGlowHelper.SetGlowByImportance(shape, importanceLevel);

        // Save the workbook with the applied glow effect
        workbook.Save("ShapeGlowByImportance.xlsx");
    }
}