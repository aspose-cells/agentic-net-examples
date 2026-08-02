// Title: C# – Fill an Excel matrix from a 2‑D array using Aspose.Cells Range
// Description: Creates a new workbook, defines a 2‑D object[,] matrix, builds a matching Range starting at A1 with Cells.CreateRange, assigns the matrix to Range.Value, and saves the file as MatrixFromArray.xlsx.
// Keywords: Aspose.Cells C# Range | fill Excel range from 2D array | populate matrix in worksheet | Cells.CreateRange example | Range.Value two dimensional array | export matrix to Excel
// Common Searches: Aspose.Cells fill range with 2D array C# | CreateRange size from array dimensions | Set Excel range value using object[,] | C# write matrix to Excel without loops | Aspose.Cells matrix layout example
// Developer Intent: Write a 2‑D array directly into an Excel range in one operation.
// Use Cases: Generate tabular reports from calculation results stored in a matrix. | Export data grids from in‑memory collections to Excel without cell‑by‑cell loops. | Build dynamic dashboards where rows and columns are determined at runtime.
// AI Prompts: Show C# code that creates a Range matching a 2D object[,] size and fills it with Range.Value using Aspose.Cells. | Explain how to change the start cell of the target range and add simple header formatting after populating the matrix. | Describe handling of mixed data types (string, int, double) when assigning a 2‑D array to Range.Value.

using System;
using Aspose.Cells;

// Creates a new workbook, defines a 2‑D object[,] matrix, builds a matching Range starting at A1 with Cells.CreateRange, assigns the matrix to Range.Value, and saves the file as MatrixFromArray.xlsx.
public class FillMatrixFromArray
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array that represents the matrix data
        object[,] matrixData = new object[,]
        {
            { "Item", "Quantity", "Price" },
            { "Apple", 10, 0.5 },
            { "Banana", 20, 0.3 },
            { "Cherry", 15, 0.8 }
        };

        // Determine the size of the array
        int rowCount = matrixData.GetLength(0);
        int columnCount = matrixData.GetLength(1);

        // Create a range that matches the size of the array (starting at cell A1)
        Aspose.Cells.Range targetRange = cells.CreateRange(0, 0, rowCount, columnCount);

        // Fill the range with the two‑dimensional array using the Range.Value property
        targetRange.Value = matrixData;

        // Save the workbook to a file
        string outputPath = "MatrixFromArray.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
