// Title: How to add thin black borders on all sides of a cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set thin black border on a single Excel cell | C# apply border style to every side of a worksheet cell with Aspose.Cells | Create thin black cell borders in an .xlsx file using Aspose.Cells API | Set cell border color and line style in Aspose.Cells for .NET | Use StyleFlag to apply only borders to a cell in Aspose.Cells C#
// Tags: Aspose.Cells thin black cell border | C# set cell border style Aspose | Excel .xlsx border formatting Aspose.Cells | StyleFlag apply borders only Aspose | CellBorders line style thin Aspose

using Aspose.Cells;
using System.Drawing;

// The example creates a workbook, accesses cell A1, defines a style with thin black borders on all four sides, applies the border style using a StyleFlag so only borders are affected, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the standard create rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access the target cell (example: A1)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample");

        // Create a style for thin black borders on all sides
        Style borderStyle = workbook.CreateStyle();

        // Set thin line style for each border
        borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

        // Set border color to black
        borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
        borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the style to the cell (only borders are affected)
        StyleFlag flag = new StyleFlag();
        flag.Borders = true;
        cell.SetStyle(borderStyle, flag);

        // Save the workbook (using the standard save rule)
        workbook.Save("output.xlsx");
    }
}
