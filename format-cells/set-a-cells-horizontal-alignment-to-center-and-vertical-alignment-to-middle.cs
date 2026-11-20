// Title: How to center text horizontally and vertically in an Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell horizontal alignment to center | C# Aspose.Cells align cell content vertically middle | center text in Excel cell using Aspose.Cells .NET API | apply both horizontal and vertical alignment to a cell with Aspose.Cells | Aspose.Cells style object set alignment center C# example
// Tags: Aspose.Cells cell alignment API C# | Excel cell style horizontal center Aspose | Excel cell style vertical middle Aspose | TextAlignmentType.Center in Aspose.Cells | C# modify cell style Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, accesses cell A1, retrieves its Style object, sets HorizontalAlignment and VerticalAlignment to TextAlignmentType.Center, reapplies the style to the cell, and saves the file as AlignedCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create rule

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the cell you want to format (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Retrieve the current style of the cell
        Style style = cell.GetStyle();

        // Set horizontal alignment to center
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Set vertical alignment to middle (center)
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook (save rule)
        workbook.Save("AlignedCell.xlsx");
    }
}
