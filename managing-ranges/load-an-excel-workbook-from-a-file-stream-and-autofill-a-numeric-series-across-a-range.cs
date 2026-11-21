// Title: C# – Load an Excel workbook from a file and autofill a linear numeric series in a column using Aspose.Cells for .NET
// Common Searches: c# aspose.cells fill column with incremental numbers | how to convert Excel cell address to row/column index with CellsHelper in Aspose.Cells | asp.net load existing Excel file and add a numeric series using Aspose.Cells | c# generate linear series from A1 to A10 with Aspose.Cells | aspose.cells autofill numeric series across a range in a worksheet
// Tags: Aspose.Cells load workbook from file | Aspose.Cells fill linear numeric series | Aspose.Cells CellsHelper address to index conversion | Aspose.Cells write values to column range | Aspose.Cells save workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example demonstrates how to load (or create) an Excel workbook, ensure the target worksheet exists, convert start and end cell addresses to row/column indexes using CellsHelper, write a linear numeric series with a specified start value and step into the defined column range, and finally save the workbook back to the same file using Aspose.Cells for .NET.
    public class ExcelSeriesFiller
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sheetName">Name of the worksheet to operate on.</param>
        /// <param name="startCell">Address of the first cell (e.g., "A1").</param>
        /// <param name="endCell">Address of the last cell in the fill range (e.g., "A10").</param>
        /// <param name="startValue">Numeric value to place in the start cell.</param>
        /// <param name="step">Increment step for the series.</param>
        public void FillSeries(string filePath, string sheetName, string startCell, string endCell, double startValue, double step)
        {
            try
            {
                // Ensure the workbook exists; create a minimal one if it does not.
                if (!File.Exists(filePath))
                {
                    Workbook newWb = new Workbook();
                    newWb.Worksheets[0].Name = sheetName;
                    newWb.Save(filePath, SaveFormat.Xlsx);
                }

                // Load the workbook from the file.
                Workbook workbook = new Workbook(filePath);

                // Get or create the target worksheet.
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    int index = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[index];
                    sheet.Name = sheetName;
                }

                // Resolve start and end cell references using CellsHelper.
                int startRow, startColumn, endRow, endColumn;
                CellsHelper.CellNameToIndex(startCell, out startRow, out startColumn);
                CellsHelper.CellNameToIndex(endCell, out endRow, out endColumn);

                // Fill the series manually (linear numeric series).
                for (int row = startRow; row <= endRow; row++)
                {
                    double value = startValue + step * (row - startRow);
                    sheet.Cells[row, startColumn].PutValue(value);
                }

                // Save the modified workbook back to the same file.
                workbook.Save(filePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error in FillSeries: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";
                string sheetName = "Sheet1";
                string startCell = "A1";
                string endCell = "A10";
                double startValue = 1;
                double step = 1;

                // Ensure the input file exists; create a new workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Workbook wb = new Workbook();
                    wb.Worksheets[0].Name = sheetName;
                    wb.Save(inputPath, SaveFormat.Xlsx);
                }

                // Fill the series in the workbook.
                ExcelSeriesFiller filler = new ExcelSeriesFiller();
                filler.FillSeries(inputPath, sheetName, startCell, endCell, startValue, step);

                // Copy the modified workbook to the output path.
                File.Copy(inputPath, outputPath, true);
                Console.WriteLine($"Series filled and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex}");
            }
        }
    }
}
