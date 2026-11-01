// Title: How to fill cells with the workbook's Accent4 theme color when they contain dates from the current month using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# highlight cells with dates from this month using theme accent color | apply workbook theme Accent4 fill to date cells in current month with Aspose.Cells .NET | C# code to color Excel cells that have dates in the current month using Aspose.Cells | retrieve Accent4 theme color and set solid background for date cells in Aspose.Cells | iterate through worksheets and style current month date cells with Aspose.Cells
// Tags: use workbook accent color fill Aspose.Cells | solid background style for date cells Aspose.Cells | obtain workbook theme color C# Aspose.Cells | loop through worksheets and style cells Aspose.Cells | conditional formatting based on month Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, gets the workbook's Accent4 theme color, scans every worksheet for cells containing DateTime values that fall within the current month, applies a solid fill using the Accent4 color to those cells, and saves the updated workbook.
class ApplyAccent4ToCurrentMonthDates
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get Accent4 color from the workbook's theme
            Color accent4Color = workbook.GetThemeColor(ThemeColorType.Accent4);

            // Current month and year
            int currentYear = DateTime.Now.Year;
            int currentMonth = DateTime.Now.Month;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check if the cell contains a date value
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            DateTime cellDate = cell.DateTimeValue;

                            // Verify the date is within the current month and year
                            if (cellDate.Year == currentYear && cellDate.Month == currentMonth)
                            {
                                // Clone the existing style
                                Style style = cell.GetStyle();

                                // Set fill pattern to solid and apply Accent4 color
                                style.Pattern = BackgroundType.Solid;
                                style.ForegroundColor = accent4Color;

                                // Apply the modified style back to the cell
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
