// Title: How to import a two‑dimensional double[,] array into an Aspose.Cells worksheet starting at row 5, column 2 in C#
// AI Prompts: Convert a double[,] matrix to an object[,] and use ImportTwoDimensionArray to place it at row index 4, column index 1. | Write C# code that imports a numeric double[,] array into an Aspose.Cells worksheet beginning at cell B5. | Generate a complete example that saves the workbook after importing the double array with a custom start position.
// Common Searches: Aspose.Cells C# import double[,] array starting at B5 | ImportTwoDimensionArray with offset row 5 column 2 example | How to place a numeric matrix into specific cells using Aspose.Cells for .NET | C# convert double[,] to object[,] for Aspose.Cells import
// Tags: ImportTwoDimensionArray start cell offset | double[,] to object[,] conversion Aspose.Cells | write numeric matrix to Excel C# | custom start row column Aspose.Cells import | Aspose.Cells workbook save after data import

using Aspose.Cells;
using System;

// Creates a workbook, converts a double[,] to object[,], imports the data into the first worksheet beginning at row 5 column 2 (cell B5) using ImportTwoDimensionArray, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a two‑dimensional double array
            double[,] data = new double[,]
            {
                { 1.1, 2.2, 3.3 },
                { 4.4, 5.5, 6.6 },
                { 7.7, 8.8, 9.9 }
            };

            // Convert double[,] to object[,] because ImportTwoDimensionArray expects object[,]
            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            object[,] objData = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    objData[i, j] = data[i, j];

            // Import the array starting at row 5 (zero‑based index 4) and column 2 (zero‑based index 1)
            sheet.Cells.ImportTwoDimensionArray(objData, 4, 1);

            // Save the workbook to a file
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
