// Title: How to apply a bold yellow style with thin borders to every cell in a named range using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells apply custom style to a named range in an existing workbook | set bold font and yellow background for cells in a named range using Aspose.Cells API | iterate over cells of a named range and assign a style with borders in .NET | retrieve named range by name and format its cells with Aspose.Cells for .NET | apply same formatting to all cells inside a defined name in Excel via Aspose.Cells
// Tags: uniform formatting for cells in a named range Aspose.Cells | bold font with yellow background and thin borders C# | access named range object Aspose.Cells | format Excel named range cells programmatically .NET | loop through range cells to set style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, creates a custom style (bold font, yellow fill, thin borders), retrieves the named range "MyRange", iterates through each cell in that range to apply the style, and saves the updated workbook.
class ApplyCustomStyleToNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string namedRangeName = "MyRange";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.IsBold = true;
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            customStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Retrieve the named range
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName(namedRangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Named range '{namedRangeName}' not found.");
                return;
            }

            // Apply the style to each cell within the range
            Worksheet ws = namedRange.Worksheet;
            int firstRow = namedRange.FirstRow;
            int firstColumn = namedRange.FirstColumn;
            int rowCount = namedRange.RowCount;
            int columnCount = namedRange.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = ws.Cells[firstRow + i, firstColumn + j];
                    cell.SetStyle(customStyle);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
