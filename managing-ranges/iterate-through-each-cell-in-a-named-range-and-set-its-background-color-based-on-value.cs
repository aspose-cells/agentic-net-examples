// Title: Iterate over a named range and apply conditional background colors to cells in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells loop through cells in a named range and change background color based on value | apply conditional background color to cells within a named range using Aspose.Cells .NET | how to set cell style programmatically for each cell in a named range Aspose.Cells example | retrieve range bounds and style cells in an Aspose.Cells workbook C# | color code positive, negative and zero values in Excel with Aspose.Cells C#
// Tags: iterate named range cells Aspose.Cells | conditional background color Aspose.Cells | style cells based on numeric value Aspose.Cells | range bounds extraction Aspose.Cells | excel cell coloring .NET Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an Excel workbook, accesses the named range "MyRange", calculates its row and column limits, and iterates through each cell. Numeric values are evaluated: positives receive a LightGreen fill, negatives a LightCoral fill, and zeros a LightGray fill. The appropriate style is applied only when needed, and the modified workbook is saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range "MyRange"
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Get the actual Range object from the named range
            AsposeRange range = namedRange.GetRange();
            if (range == null)
            {
                Console.WriteLine("Unable to resolve the range for 'MyRange'.");
                return;
            }

            // Determine the bounds of the range
            int firstRow = range.FirstRow;
            int lastRow = firstRow + range.RowCount - 1;
            int firstCol = range.FirstColumn;
            int lastCol = firstCol + range.ColumnCount - 1;

            // Iterate through each cell in the range
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    object val = cell.Value;
                    Color bgColor = Color.White; // default

                    // Set background color based on numeric value
                    if (val is double d)
                    {
                        if (d > 0)
                            bgColor = Color.LightGreen;
                        else if (d < 0)
                            bgColor = Color.LightCoral;
                        else
                            bgColor = Color.LightGray;
                    }

                    // Apply the background color if needed
                    if (bgColor != Color.White)
                    {
                        Style style = cell.GetStyle();
                        style.ForegroundColor = bgColor;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);
                    }
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
