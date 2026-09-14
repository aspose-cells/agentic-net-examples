// Title: Color cells with dates from the current month using the workbook’s Accent4 theme color in Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds cells containing DateTime values matching the current month and year, and applies a solid Accent4 background fill. | Create a reusable method in Aspose.Cells to highlight date cells of the current month by setting the cell style’s ForegroundColor to the Accent4 theme color and saving the workbook.
// Common Searches: how to highlight Excel cells that contain dates from this month using Aspose.Cells C# | apply theme accent4 fill to date cells in a .NET workbook | Aspose.Cells iterate used range and set background color based on DateTime value | C# change cell style for current month dates in an existing XLSX file | set solid fill color for cells with dates in the current month using Aspose.Cells API
// Tags: accent4 theme color cell style Aspose.Cells | date cell background fill month-year .NET | iterate worksheet used range Aspose.Cells | assign background pattern to DateTime cells C# | apply foreground color based on month-year Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, iterates through each worksheet's used range, checks each cell for a DateTime value, and if the date falls within the current month and year, applies a solid fill using the workbook’s Accent4 theme color before saving the modified file.
class ApplyAccent4ToCurrentMonthDates
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Use a typical Accent4 theme color (blue)
        Color accent4Color = Color.FromArgb(0, 112, 192);

        // Current month and year
        DateTime now = DateTime.Now;
        int currentMonth = now.Month;
        int currentYear = now.Year;

        try
        {
            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Iterate through cells in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check for DateTime values
                        if (cell.Type == CellValueType.IsDateTime && cell.Value is DateTime cellDate)
                        {
                            // Apply style if the date is in the current month/year
                            if (cellDate.Month == currentMonth && cellDate.Year == currentYear)
                            {
                                Style style = cell.GetStyle();
                                style.ForegroundColor = accent4Color;
                                style.Pattern = BackgroundType.Solid;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
