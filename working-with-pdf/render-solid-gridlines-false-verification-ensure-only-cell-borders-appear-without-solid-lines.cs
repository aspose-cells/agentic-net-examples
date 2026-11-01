// Title: How to hide Excel gridlines and add thin black borders to a specific cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide worksheet gridlines and add thin borders to a range | How to turn off default Excel gridlines while applying custom cell borders with Aspose.Cells | C# example for disabling solid gridlines and setting black borders on cells A1:C3 using Aspose.Cells | Remove Excel gridlines but keep cell borders in a .NET workbook with Aspose.Cells | Aspose.Cells hide gridlines and style cell borders programmatically in C#
// Tags: hide worksheet gridlines Aspose.Cells | apply thin black borders to cell range Aspose.Cells | set cell border style C# Aspose.Cells | disable solid gridlines Excel Aspose.Cells | save workbook with custom borders Aspose.Cells Xlsx

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, disables the default gridlines, defines the A1:C3 range, applies thin black borders to each side of every cell in that range, and saves the file as GridlinesFalse.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Hide the default solid gridlines
        sheet.IsGridlinesVisible = false;

        // Define the range where borders will be applied (e.g., A1:C3)
        CellArea area = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = 2,
            EndColumn = 2
        };

        // Apply thin black borders to each cell in the defined range
        for (int row = area.StartRow; row <= area.EndRow; row++)
        {
            for (int col = area.StartColumn; col <= area.EndColumn; col++)
            {
                Style style = workbook.CreateStyle();

                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                style.Borders[BorderType.TopBorder].Color = Color.Black;
                style.Borders[BorderType.BottomBorder].Color = Color.Black;
                style.Borders[BorderType.LeftBorder].Color = Color.Black;
                style.Borders[BorderType.RightBorder].Color = Color.Black;

                sheet.Cells[row, col].SetStyle(style);
            }
        }

        // Save the workbook
        workbook.Save("GridlinesFalse.xlsx", SaveFormat.Xlsx);
    }
}
