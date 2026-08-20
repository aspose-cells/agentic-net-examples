// Title: Apply Dense Diagonal Crosshatch Fill to Header Cells with Aspose.Cells for .NET
// Description: This example creates a new Workbook, defines a style using the ThickDiagonalCrosshatch pattern with DarkBlue hatch lines on a LightYellow background, applies the style to the first‑row header cells (A1‑D1), and saves the result as HeaderPatternDemo.xlsx.
// Keywords: Aspose.Cells | C# fill pattern | BackgroundType.ThickDiagonalCrosshatch | header cell style | Excel cell pattern | custom hatch fill | .NET Excel styling | dense hatch pattern
// Common Searches: Aspose.Cells set thick diagonal crosshatch pattern | C# apply hatch fill to Excel header row | BackgroundType.ThickDiagonalCrosshatch example | How to style Excel header with pattern using Aspose.Cells | Change foreground and background colors for hatch fill Aspose.Cells
// Developer Intent: Create a dense hatch style and apply it to header cells in an Excel worksheet using Aspose.Cells.
// Use Cases: Emphasize header rows in financial statements with a distinctive pattern | Generate printable reports where headers need visual distinction | Automate consistent patterned styling across multiple worksheets for corporate branding | Design dashboards with patterned headers to improve visual hierarchy
// AI Prompts: Generate a reusable method that applies any BackgroundType pattern with configurable foreground and background colors to a specified cell range in Aspose.Cells. | Show how to export the styled workbook to PDF while preserving the hatch fill pattern. | Provide code to replace the dense hatch with a light diagonal hatch without altering other style attributes. | Explain performance considerations when applying styles to large ranges in Aspose.Cells.

using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderPatternDemo
{
    // This example creates a new Workbook, defines a style using the ThickDiagonalCrosshatch pattern with DarkBlue hatch lines on a LightYellow background, applies the style to the first‑row header cells (A1‑D1), and saves the result as HeaderPatternDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style with a dense hatch (thick diagonal crosshatch) pattern
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.ThickDiagonalCrosshatch; // dense hatch
            headerStyle.ForegroundColor = Color.DarkBlue;   // color of the hatch lines
            headerStyle.BackgroundColor = Color.LightYellow; // background color behind the hatch

            // Apply the style to the header row (e.g., first row A1 to D1)
            for (int col = 0; col < 4; col++)
            {
                Cell cell = sheet.Cells[0, col];
                cell.PutValue($"Header {col + 1}");
                cell.SetStyle(headerStyle);
            }

            // Save the workbook
            workbook.Save("HeaderPatternDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
