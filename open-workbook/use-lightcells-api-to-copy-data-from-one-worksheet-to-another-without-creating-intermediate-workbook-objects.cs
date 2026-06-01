using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsCopyDemo
{
    class Program
    {
        static void Main()
        {
            const string sourcePath = "Source.xlsx";
            const string destPath = "Destination.xlsx";

            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                    throw new FileNotFoundException($"Source file not found: {sourcePath}");

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // first worksheet

                // Create an empty destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0]; // first worksheet

                // Copy data from source to destination using standard Cells API
                Cells srcCells = sourceSheet.Cells;
                Cells dstCells = destSheet.Cells;

                int maxRow = srcCells.MaxDataRow;
                int maxCol = srcCells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get the source cell
                        Cell srcCell = srcCells[row, col];
                        if (srcCell != null && srcCell.Type != CellValueType.IsNull)
                        {
                            // Copy value and style to destination cell
                            Cell dstCell = dstCells[row, col];
                            dstCell.PutValue(srcCell.Value);
                            dstCell.SetStyle(srcCell.GetStyle());
                        }
                    }
                }

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Workbook copied successfully to '{destPath}'.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}