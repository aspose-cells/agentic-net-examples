// Title: Loop through a named range and apply background colors based on numeric values with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate over cells in a named range and change background color | set cell background color according to numeric value in a named range using Aspose.Cells | programmatically apply conditional formatting to a named range with Aspose.Cells .NET | retrieve a named range and style each cell based on its value in C# Aspose.Cells | how to color cells green, red, yellow based on value in Aspose.Cells workbook
// Tags: named range cell iteration Aspose.Cells | apply solid background pattern C# Aspose.Cells | conditional background color based on numeric value | Aspose.Cells GetRange method usage | style cells programmatically in Excel with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the named range "MyRange", loops through every cell in that range, determines if the cell contains a numeric value and assigns LightGreen, LightCoral, LightYellow, or LightGray as the background color, then saves the modified file as "output.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Resolve the named range to a Range object (use fully qualified name to avoid ambiguity)
            Aspose.Cells.Range range = namedRange.GetRange();
            if (range == null)
            {
                Console.WriteLine("Unable to resolve the named range to a range.");
                return;
            }

            Worksheet ws = range.Worksheet;
            int startRow = range.FirstRow;
            int endRow = startRow + range.RowCount - 1;
            int startCol = range.FirstColumn;
            int endCol = startCol + range.ColumnCount - 1;

            // Iterate through each cell in the range and set background color based on value
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = ws.Cells[row, col];

                    Color bgColor = Color.LightGray; // default for non-numeric
                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double numericValue))
                    {
                        if (numericValue > 0)
                            bgColor = Color.LightGreen;
                        else if (numericValue < 0)
                            bgColor = Color.LightCoral;
                        else
                            bgColor = Color.LightYellow;
                    }

                    Style style = cell.GetStyle();
                    style.ForegroundColor = bgColor;
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
