// Title: Apply a diagonal crosshatch fill pattern to cells containing {{placeholder}} text using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, scans each cell for the "{{" and "}}" markers, creates a Style with BackgroundType.DiagonalCrosshatch, sets LightGray as the foreground color and White as the background color, applies the style to the matching cells, and saves the workbook. | Show a step‑by‑step example of using Aspose.Cells to detect placeholder tokens in a worksheet and assign a diagonal crosshatch background pattern only to those cells.
// Common Searches: aspocells set diagonal crosshatch pattern for cells with {{placeholder}} tokens | c# aspocells apply custom fill pattern based on cell text content | how to style excel cells containing placeholder markers using Aspose.Cells | aspocells conditional formatting with diagonal crosshatch fill in .NET | c# code to add diagonal crosshatch background to specific Excel cells
// Tags: aspocells diagonal crosshatch cell style | c# conditional style for placeholder cells | backgroundtype.diagonalcrosshatch usage | apply custom fill pattern to excel cells c# | detect {{placeholder}} tokens with aspocells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, creates a Style with BackgroundType.DiagonalCrosshatch, LightGray foreground and White background, scans the used range for cells containing "{{...}}" placeholders, applies the style to those cells, and saves the modified file.
class ApplyDiagonalCrosshatchFill
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style with diagonal crosshatch fill pattern
            Style crosshatchStyle = workbook.CreateStyle();
            crosshatchStyle.Pattern = BackgroundType.DiagonalCrosshatch; // fill pattern
            crosshatchStyle.ForegroundColor = Color.LightGray;           // pattern color
            crosshatchStyle.BackgroundColor = Color.White;              // cell background

            // Determine the used range of the worksheet (use Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            // Apply the style to cells containing placeholders like {{placeholder}}
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    string text = cell.StringValue;

                    if (!string.IsNullOrEmpty(text) && text.Contains("{{") && text.Contains("}}"))
                    {
                        cell.SetStyle(crosshatchStyle);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
