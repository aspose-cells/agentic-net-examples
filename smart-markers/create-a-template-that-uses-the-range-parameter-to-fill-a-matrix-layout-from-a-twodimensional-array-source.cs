using System;
using Aspose.Cells;

class FillMatrixFromArray
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Two‑dimensional array source (matrix)
            int[,] matrix = new int[,]
            {
                { 1, 2, 3, 4 },
                { 5, 6, 7, 8 },
                { 9, 10, 11, 12 }
            };

            // Determine the size of the matrix
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            // Create a range that matches the matrix size, starting at cell A1
            // Use fully qualified name to avoid conflict with System.Range
            Aspose.Cells.Range targetRange = cells.CreateRange(0, 0, rows, cols);

            // Fill the range with the matrix using the Range.Value property
            targetRange.Value = matrix;

            // Save the workbook
            workbook.Save("MatrixFromArray.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}