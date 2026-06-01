using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyRangeWithSizes
{
    class Program
    {
        static void Main()
        {
            try
            {
                string sourcePath = "source.xlsx";
                string destPath = "dest.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define the source range to copy (A1:C10)
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1:C10");

                // Create a new (empty) destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the destination range (same size as source)
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:C10");

                // Copy data, formulas, formatting, etc.
                destRange.Copy(sourceRange);

                // Preserve column widths (in characters)
                int firstCol = sourceRange.FirstColumn;
                int totalCols = sourceRange.ColumnCount;
                for (int i = 0; i < totalCols; i++)
                {
                    int colIndex = firstCol + i;
                    double width = sourceSheet.Cells.GetColumnWidth(colIndex);
                    destSheet.Cells.SetColumnWidth(colIndex, width);
                }

                // Preserve row heights (in points)
                int firstRow = sourceRange.FirstRow;
                int totalRows = sourceRange.RowCount;
                for (int i = 0; i < totalRows; i++)
                {
                    int rowIndex = firstRow + i;
                    double height = sourceSheet.Cells.GetRowHeight(rowIndex);
                    destSheet.Cells.SetRowHeight(rowIndex, height);
                }

                // Ensure destination directory exists
                string destDir = Path.GetDirectoryName(Path.GetFullPath(destPath));
                if (!Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Destination workbook saved to: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}