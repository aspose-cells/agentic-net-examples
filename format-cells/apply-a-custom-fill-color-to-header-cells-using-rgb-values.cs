// Title: How to apply a custom RGB fill color to header cells in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set solid background color for a specific row using RGB values | Apply custom fill color to header range A1:D1 with Aspose.Cells .NET | Create and apply a style with Color.FromArgb to Excel header cells in C# | Use StyleFlag and BackgroundType.Solid to color an Excel header row in C# | Set light‑blue background for header cells using Aspose.Cells style flag
// Tags: rgb fill style Aspose.Cells | excel header style C# | StyleFlag ApplyStyle Excel range | BackgroundType.Solid custom color | Color.FromArgb cell style .NET

using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, defines the header range A1:D1, builds a style with a solid light‑blue fill using Color.FromArgb(173,216,230), applies the style to the range via a StyleFlag, adds header text, and saves the file as HeaderStyled.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the header range (e.g., cells A1 to D1)
        int startRow = 0;      // Row index is zero‑based
        int startColumn = 0;   // Column A
        int endRow = 0;
        int endColumn = 3;     // Column D

        // Create a style for the header cells
        Style headerStyle = workbook.CreateStyle();

        // Set a custom fill color using RGB values (example: light blue)
        headerStyle.ForegroundColor = Color.FromArgb(173, 216, 230);
        headerStyle.Pattern = BackgroundType.Solid;

        // Apply the style to the defined range
        StyleFlag flag = new StyleFlag { All = true };
        sheet.Cells.CreateRange(startRow, startColumn,
                                endRow - startRow + 1,
                                endColumn - startColumn + 1).ApplyStyle(headerStyle, flag);

        // Optional: add header text
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["C1"].PutValue("Header 3");
        sheet.Cells["D1"].PutValue("Header 4");

        // Save the workbook
        workbook.Save("HeaderStyled.xlsx");
    }
}
