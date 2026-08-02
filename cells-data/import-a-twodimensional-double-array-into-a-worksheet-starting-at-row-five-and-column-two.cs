// Title: Import a 2D double[,] array into an Aspose.Cells worksheet at row 5, column 2 (C#)
// Description: Creates a double[,] matrix, defines zero‑based start indices (row 4, column 1) for Excel row 5/column 2, loops through the matrix and writes each value with Cells[row, col].PutValue, then saves the workbook as TwoDimArrayImport.xlsx. The example also notes that Cells.ImportArray can perform the same operation in a single call.
// Keywords: Aspose.Cells C# import double array | double[,] to Excel | write 2D array Aspose.Cells | Cells.PutValue loop | Cells.ImportArray example | start row column Aspose.Cells | Excel row 5 column 2 | C# Excel matrix export | Aspose.Cells worksheet data import | GitHub Aspose.Cells sample
// Common Searches: How to import a 2D double array into Aspose.Cells C# | Aspose.Cells putvalue starting at row 5 column 2 | Write double[,] data to Excel worksheet with Aspose.Cells | Cells.ImportArray with custom start position | C# export numeric matrix to Excel using Aspose.Cells
// Developer Intent: Write the contents of a two‑dimensional double[,] matrix into a worksheet beginning at Excel row 5, column 2.
// Use Cases: Populate a financial model where the matrix must start after header rows. | Export simulation results stored in a double[,] matrix to a predefined area of an Excel report. | Generate data for a chart by writing a numeric matrix to a worksheet with a fixed offset.
// AI Prompts: Show a C# example that imports a 2D double[,] into an Aspose.Cells worksheet at a given start row and column using a loop. | Provide code that uses Cells.ImportArray to place a double[,] matrix starting at row 5, column 2. | Explain how to adapt the sample for variable start positions and large matrices while keeping performance optimal.

using System;
using Aspose.Cells;

// Creates a double[,] matrix, defines zero‑based start indices (row 4, column 1) for Excel row 5/column 2, loops through the matrix and writes each value with Cells[row, col].PutValue, then saves the workbook as TwoDimArrayImport.xlsx. The example also notes that Cells.ImportArray can perform the same operation in a single call.
class Program
{
    static void Main()
    {
        // Sample two‑dimensional double array
        double[,] data = new double[,]
        {
            { 1.1, 2.2, 3.3 },
            { 4.4, 5.5, 6.6 },
            { 7.7, 8.8, 9.9 }
        };

        // Starting position: row 5 (zero‑based index 4), column 2 (zero‑based index 1)
        int startRow = 4;
        int startCol = 1;

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the array values cell by cell
        int rows = data.GetLength(0);
        int cols = data.GetLength(1);
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                cells[startRow + i, startCol + j].PutValue(data[i, j]);
            }
        }

        // Save the workbook
        workbook.Save("TwoDimArrayImport.xlsx");
    }
}
