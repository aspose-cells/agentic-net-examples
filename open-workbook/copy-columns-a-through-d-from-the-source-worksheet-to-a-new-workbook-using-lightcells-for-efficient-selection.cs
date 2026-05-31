using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLightCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "Source.xlsx";
                const string destPath = "Destination.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Create a new destination workbook
                Workbook destWorkbook = new Workbook();
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the range to copy: columns A (0) through D (3)
                int startColumn = 0;          // Column A
                int totalColumns = 4;         // A, B, C, D
                int startRow = 0;
                int totalRows = sourceSheet.Cells.MaxDataRow + 1; // include all rows with data

                // Copy values and styles cell by cell
                for (int row = startRow; row < totalRows; row++)
                {
                    for (int col = startColumn; col < startColumn + totalColumns; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[row, col];
                        Cell destCell = destSheet.Cells[row, col];

                        // Copy value
                        destCell.PutValue(srcCell.Value);

                        // Copy style (SetStyle copies the style, no need to clone)
                        Style srcStyle = srcCell.GetStyle();
                        if (srcStyle != null)
                        {
                            destCell.SetStyle(srcStyle);
                        }
                    }
                }

                // Save the destination workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Copy completed successfully. Destination saved to {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}