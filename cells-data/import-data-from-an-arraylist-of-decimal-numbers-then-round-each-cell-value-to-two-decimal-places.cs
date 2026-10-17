// Title: Import an ArrayList of decimal numbers into an Excel worksheet and round each cell to two decimal places using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an ArrayList of decimal values, writes them to column A of a new workbook with Aspose.Cells, then rounds every numeric cell in the used range to two decimal places before saving as XLSX. | Show how to use Aspose.Cells' MaxDisplayRange to iterate over all populated cells and apply Math.Round with MidpointRounding.AwayFromZero to each numeric cell after importing data from a collection. | Create a .NET example that converts each item in an ArrayList to decimal, inserts it into a worksheet, formats the cells to display two decimal places, and saves the workbook.
// Common Searches: C# Aspose.Cells load decimal numbers from a List into Excel column A | Apply two‑decimal rounding to all numeric cells after populating a worksheet with Aspose.Cells | Iterate over used range in Aspose.Cells and round values in C# | Save Aspose.Cells workbook as XLSX after rounding decimal data | Add decimal values from an ArrayList to cells in Aspose.Cells
// Tags: insert decimal data into worksheet Aspose.Cells | round numeric cells to two decimals Aspose.Cells | enumerate worksheet cells Aspose.Cells | write workbook to XLSX format Aspose.Cells

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, imports a list of decimal numbers from an ArrayList into column A, iterates over the used range to round each numeric cell to two decimal places, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: ArrayList of decimal numbers.
            ArrayList data = new ArrayList
            {
                12.3456m,
                78.9012m,
                3.14159m,
                0.9999m,
                123.4567m
            };

            // Import data into column A starting at row 0.
            for (int i = 0; i < data.Count; i++)
            {
                // Safely convert each item to decimal and assign to the cell.
                decimal value = Convert.ToDecimal(data[i]);
                sheet.Cells[i, 0].PutValue(value);
            }

            // Round each numeric cell value in the used range to two decimal places.
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
            {
                for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double original = cell.DoubleValue;
                        double rounded = Math.Round(original, 2, MidpointRounding.AwayFromZero);
                        cell.PutValue(rounded);
                    }
                }
            }

            // Ensure the output directory exists.
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
