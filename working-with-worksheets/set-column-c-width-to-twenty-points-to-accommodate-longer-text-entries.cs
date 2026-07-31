// Title: Set Column C Width to 20 pts in Excel with Aspose.Cells for .NET
// Description: Demonstrates how to convert 20 points to inches and apply the measurement to column C (index 2) using Aspose.Cells' SetColumnWidthInch method, then inserts a long string to verify the layout and saves the file as ColumnCWidth.xlsx.
// Keywords: Aspose.Cells column width points | C# SetColumnWidthInch | Excel column width inches .NET | adjust column C size Aspose.Cells | Excel automation column sizing
// Common Searches: Aspose.Cells set column width in points C# | How to use SetColumnWidthInch for column C | Convert points to inches for Excel column width | C# code to set Excel column width to 20 pts
// Developer Intent: Apply a 20‑point (≈0.278 in) width to column C of an Excel worksheet via Aspose.Cells in a C# project.
// Use Cases: Designing a report where column C must hold lengthy text without wrapping. | Creating a template with a fixed column dimension to preserve layout consistency. | Exporting data that must match a design specification defined in typographic points.
// AI Prompts: Generate C# code that sets column D to 30 pts using Aspose.Cells. | Explain the relationship between inches, points, and Excel column width units in Aspose.Cells. | Show how to set column width in pixels instead of points with Aspose.Cells.

using Aspose.Cells;

// Demonstrates how to convert 20 points to inches and apply the measurement to column C (index 2) using Aspose.Cells' SetColumnWidthInch method, then inserts a long string to verify the layout and saves the file as ColumnCWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Column C has zero‑based index 2.
        // 1 point = 1/72 inch, so 20 points = 20/72 inches.
        double inches = 20.0 / 72.0;

        // Set the width of column C using the SetColumnWidthInch method.
        cells.SetColumnWidthInch(2, inches);

        // Example data to demonstrate the width
        cells["C1"].PutValue("This is a longer text entry that fits within 20 points.");

        // Save the workbook
        workbook.Save("ColumnCWidth.xlsx");
    }
}
