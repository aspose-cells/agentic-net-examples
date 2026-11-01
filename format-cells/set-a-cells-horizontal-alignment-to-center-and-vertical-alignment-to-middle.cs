// Title: Center cell text horizontally and vertically using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell horizontal alignment to center | C# Aspose.Cells vertical alignment middle for a specific cell | How to apply centered alignment to an Excel cell with Aspose.Cells API | Center text in Excel cell programmatically using Aspose.Cells .NET | Set both horizontal and vertical alignment of a cell in Aspose.Cells C#
// Tags: Aspose.Cells set cell alignment | center alignment for Excel cell Aspose.Cells | C# Aspose.Cells style horizontal vertical alignment | apply cell style Aspose.Cells .NET | Excel cell formatting Aspose.Cells C#

using Aspose.Cells;

// Creates a new workbook, accesses cell A1, retrieves its style, sets HorizontalAlignment and VerticalAlignment to Center, reapplies the style, and saves the file as AlignedCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the target cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Retrieve the current style of the cell
        Style style = cell.GetStyle();

        // Set horizontal alignment to center
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Set vertical alignment to middle (center)
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("AlignedCell.xlsx");
    }
}
