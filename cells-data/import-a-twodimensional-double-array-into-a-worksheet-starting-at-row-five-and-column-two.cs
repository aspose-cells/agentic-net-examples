// Title: C# example: Import a two‑dimensional double[,] array into an Aspose.Cells worksheet beginning at cell B5 (row 5, column 2)
// AI Prompts: Generate C# code that uses Aspose.Cells to write each element of a double[,] array into a worksheet starting at cell B5 and then saves the workbook. | Create a nested‑loop snippet that populates an Excel sheet with values from a 2D double array, beginning at row 5 column 2, using Cells.PutValue in Aspose.Cells.
// Common Searches: C# Aspose.Cells write 2D double array to specific cell range B5 | How to use Cells.PutValue to import a double[,] starting at row 5 column 2 in Aspose.Cells | Aspose.Cells populate worksheet from a multidimensional double array beginning at B5 | Importing a two‑dimensional double matrix into Excel with Aspose.Cells C# example
// Tags: import 2d double array Aspose.Cells C# | write double[,] to Excel cell range Aspose.Cells | populate worksheet starting at B5 Aspose.Cells | Cells.PutValue multidimensional array example | C# Aspose.Cells write data to specific row and column

using System;
using Aspose.Cells;

// The sample creates a Workbook, defines a double[,] matrix, and uses a nested loop with Cells.PutValue to insert each value into the first worksheet starting at row 5, column 2 (cell B5). The workbook is then saved as TwoDimensionalArrayImport.xlsx.
class ImportTwoDimensionalArrayDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample two‑dimensional double array
        double[,] data = new double[,]
        {
            { 1.23, 4.56, 7.89 },
            { 10.11, 12.13, 14.15 },
            { 16.17, 18.19, 20.21 }
        };

        // Starting position: row 5 (index 4), column 2 (index 1) – zero‑based indices
        int startRow = 4;   // corresponds to Excel row 5
        int startColumn = 1; // corresponds to Excel column B

        // Import the array by iterating through its dimensions
        int rows = data.GetLength(0);
        int cols = data.GetLength(1);
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                cells[startRow + i, startColumn + j].PutValue(data[i, j]);
            }
        }

        // Save the workbook
        workbook.Save("TwoDimensionalArrayImport.xlsx");
    }
}
