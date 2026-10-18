// Title: Highlight Excel error cells with the workbook's Accent3 theme color using Aspose.Cells for .NET
// Common Searches: how to color cells that contain error values with a theme accent in Aspose.Cells C# | apply workbook Accent3 theme as background fill for error cells in a .NET Excel file | using GetThemeColor to style cells where CellValueType.IsError is true with Aspose.Cells | C# code to set solid background for Excel cells that have error values | iterate used range and change background of error cells in Aspose.Cells workbook
// Tags: apply accent theme color to error cells Aspose.Cells | GetThemeColor API for conditional styling C# | set solid background fill for CellValueType.IsError | iterate worksheet used range Aspose.Cells | format Excel error values with theme color .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, retrieves the Accent3 theme color via GetThemeColor, walks through each worksheet's used range, detects cells whose type is CellValueType.IsError, and applies a solid background fill using the Accent3 color before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the Accent3 theme color (compatible across versions)
                Color accent3Color = workbook.GetThemeColor(ThemeColorType.Accent3);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the sheet
                    var usedRange = sheet.Cells.MaxDisplayRange; // Aspose.Cells.Range

                    // Loop through each cell in the used range
                    for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
                    {
                        for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Check if the cell contains an error value
                            if (cell.Type == CellValueType.IsError)
                            {
                                // Apply the Accent3 color as the background fill
                                Style style = cell.GetStyle();
                                style.ForegroundColor = accent3Color;
                                style.Pattern = BackgroundType.Solid;
                                cell.SetStyle(style);
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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
