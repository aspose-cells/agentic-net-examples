// Title: C# – Apply Diagonal Crosshatch Fill to Placeholder Cells with Aspose.Cells
// Description: Shows how to create a workbook, insert placeholder tokens (e.g., {{Name}}, {{Date}}), scan the used range, and apply a diagonal‑crosshatch background (light gray on white) to those cells using Aspose.Cells for .NET, then save the file as an .xlsx.
// Keywords: Aspose.Cells | C# cell background pattern | diagonal crosshatch fill | placeholder highlighting | Excel template styling | BackgroundType.DiagonalCrosshatch | Aspose.Cells .NET | programmatic cell style
// Common Searches: Aspose.Cells set diagonal crosshatch pattern | highlight placeholder cells Aspose.Cells C# | apply custom fill pattern to cells Aspose | detect {{}} tokens in Excel with Aspose | change cell background color Aspose.Cells .NET
// Developer Intent: Programmatically highlight cells that contain {{…}} placeholders by applying a diagonal‑crosshatch fill.
// Use Cases: Visually distinguish template placeholders before generating reports. | Guide end‑users to edit placeholder cells in an Excel template. | Mark placeholder cells for downstream processing or validation.
// AI Prompts: Generate C# code using Aspose.Cells that finds cells with {{token}} placeholders and applies a diagonal‑crosshatch background pattern. | Provide an example that iterates over a worksheet, detects placeholder text, and sets both foreground and background colors for a custom fill style. | Explain how to change the pattern type or colors for highlighting placeholders in an existing Aspose.Cells workbook.

using System.Drawing;
using Aspose.Cells;

// Shows how to create a workbook, insert placeholder tokens (e.g., {{Name}}, {{Date}}), scan the used range, and apply a diagonal‑crosshatch background (light gray on white) to those cells using Aspose.Cells for .NET, then save the file as an .xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample cells containing placeholder text
        sheet.Cells["A1"].PutValue("{{Name}}");
        sheet.Cells["B2"].PutValue("{{Date}}");
        sheet.Cells["C3"].PutValue("Regular Text");

        // Define placeholder delimiters
        const string placeholderStart = "{{";
        const string placeholderEnd = "}}";

        // Iterate over the used range of the worksheet
        Cells cells = sheet.Cells;
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue;
                    if (text.Contains(placeholderStart) && text.Contains(placeholderEnd))
                    {
                        // Retrieve the cell's current style
                        Style style = cell.GetStyle();

                        // Apply diagonal crosshatch pattern
                        style.Pattern = BackgroundType.DiagonalCrosshatch;

                        // Set foreground and background colors for the pattern
                        style.ForegroundColor = Color.LightGray;
                        style.BackgroundColor = Color.White;

                        // Apply the modified style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // Save the workbook
        workbook.Save("PlaceholderPattern.xlsx", SaveFormat.Xlsx);
    }
}
