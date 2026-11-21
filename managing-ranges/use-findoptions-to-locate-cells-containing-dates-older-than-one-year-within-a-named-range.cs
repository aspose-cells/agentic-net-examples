// Title: Highlight dates older than one year in a named range using FindOptions with Aspose.Cells for .NET
// Common Searches: Aspose.Cells FindOptions example to locate dates older than one year in a named range | C# code to highlight cells with dates older than a year using Aspose.Cells | How to filter date values older than one year within a named range in Aspose.Cells .NET | Find and style old dates in Excel workbook using Aspose.Cells FindOptions C#
// Tags: findoptions date filtering aspnet | named range cell iteration aspnet | excel cell style based on date aspnet | c# highlight old dates aspnet | aspnet excel workbook date processing

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, retrieves the named range "MyRange", iterates through each cell in that range, identifies DateTime values older than one year, applies a yellow background style to those cells, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "MyRange"
            // Use an alias to avoid conflict with System.Range
            Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"MyRange\" not found.");
                return;
            }

            // Determine the bounds of the named range
            int startRow = namedRange.FirstRow;
            int startCol = namedRange.FirstColumn;
            int endRow = startRow + namedRange.RowCount - 1;
            int endCol = startCol + namedRange.ColumnCount - 1;

            // Iterate through each cell in the range and highlight dates older than one year
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    try
                    {
                        Cell cell = namedRange.Worksheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsDateTime)
                        {
                            DateTime cellDate = cell.DateTimeValue;
                            if (cellDate < DateTime.Now.AddYears(-1))
                            {
                                Style style = cell.GetStyle();
                                style.ForegroundColor = Color.Yellow;
                                style.Pattern = BackgroundType.Solid;
                                cell.SetStyle(style);
                            }
                        }
                    }
                    catch (Exception cellEx)
                    {
                        Console.WriteLine($"Error processing cell [{row}, {col}]: {cellEx.Message}");
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
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
