// Title: C# – Import a 2‑D double[,] array into Aspose.Cells starting at row 5, column 2
// Description: Shows how to create a Workbook, extract each row from a double[,] matrix, and write it horizontally into a worksheet with Cells.ImportArray, beginning at the fifth row (index 4) and second column (index 1), then save as TwoDimensionalArrayImport.xlsx.
// Keywords: Aspose.Cells | C# import double array | Cells.ImportArray | double[,] to worksheet | 2D array Excel | start row column | Aspose.Cells C# example | import matrix Aspose | Excel automation .NET | write numeric array
// Common Searches: Aspose.Cells import double[,] array C# | How to write a 2D double array to Excel with Aspose | Cells.ImportArray start at specific cell | C# write matrix to Excel starting at row 5 column 2 | Aspose.Cells place data after header rows
// Developer Intent: Insert a two‑dimensional double[,] matrix into a worksheet at a defined offset (row 5, column 2) using Aspose.Cells.
// Use Cases: Add sensor measurements after four header rows in a template. | Populate a financial matrix in a pre‑formatted report without overwriting titles. | Load simulation results into an existing workbook, aligning with layout by offsetting the start cell. | Export statistical tables into a workbook that already contains charts and headings. | Insert data for a chart series beginning at a specific cell range.
// AI Prompts: Generate C# code that uses Aspose.Cells to import a double[,] array into a worksheet starting at row 5, column 2. | Explain the Cells.ImportArray parameters for horizontal versus vertical import. | Modify the example to import the array vertically while keeping the same start cell. | Show how to import a jagged double[] array instead of a 2D array with Aspose.Cells. | Provide code that imports the array and applies numeric formatting to the cells.

using System;
using Aspose.Cells;

// Shows how to create a Workbook, extract each row from a double[,] matrix, and write it horizontally into a worksheet with Cells.ImportArray, beginning at the fifth row (index 4) and second column (index 1), then save as TwoDimensionalArrayImport.xlsx.
class ImportTwoDimensionalDoubleArray
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample two‑dimensional double array (5 rows × 3 columns)
        double[,] data = new double[,]
        {
            { 1.1, 2.2, 3.3 },
            { 4.4, 5.5, 6.6 },
            { 7.7, 8.8, 9.9 },
            { 10.1, 11.2, 12.3 },
            { 13.4, 14.5, 15.6 }
        };

        // Define the starting position: row 5 (index 4), column 2 (index 1)
        int startRow = 4;    // zero‑based index for the 5th row
        int startColumn = 1; // zero‑based index for the 2nd column

        // Import each row of the 2‑D array horizontally
        int totalRows = data.GetLength(0);
        int totalCols = data.GetLength(1);
        for (int r = 0; r < totalRows; r++)
        {
            // Extract a single row into a one‑dimensional double array
            double[] rowValues = new double[totalCols];
            for (int c = 0; c < totalCols; c++)
            {
                rowValues[c] = data[r, c];
            }

            // Import the row horizontally (isVertical = false)
            cells.ImportArray(rowValues, startRow + r, startColumn, false);
        }

        // Save the workbook to a file
        workbook.Save("TwoDimensionalArrayImport.xlsx");
    }
}
