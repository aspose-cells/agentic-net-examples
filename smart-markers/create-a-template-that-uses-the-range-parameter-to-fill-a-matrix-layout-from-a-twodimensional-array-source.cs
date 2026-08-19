// Title: Fill an Excel matrix from a 2‑D array using Aspose.Cells Range in C#
// Description: Creates a new Workbook, builds a two‑dimensional int[,] matrix, computes its dimensions, generates a matching Range with Cells.CreateRange starting at A1, assigns the matrix to Range.Value, and saves the file as MatrixFromArray.xlsx.
// Keywords: Aspose.Cells C# Range.Value | populate Excel from 2D array | CreateRange matrix size | write multi‑dimensional array to worksheet | Excel matrix fill example
// Common Searches: Aspose.Cells fill range with 2D int array C# | CreateRange for matrix dimensions Aspose.Cells | Assign two‑dimensional array to Excel range | Save workbook after setting Range.Value | C# write matrix to Excel using Aspose
// Developer Intent: Generate a worksheet range that matches a matrix’s dimensions and populate it in a single operation by assigning a two‑dimensional array to Range.Value.
// Use Cases: Export calculation results stored in a 2‑D array directly to an Excel table without cell‑by‑cell loops. | Create a heat‑map or score grid by writing numeric matrix data to a worksheet range in one step. | Produce financial or statistical reports where data is already available as a multi‑dimensional array.
// AI Prompts: Show how to start the matrix at cell B2 instead of A1. | Demonstrate writing a string[,] array to a range with Aspose.Cells. | Explain how to apply borders and background colors to the range after assigning the array.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Creates a new Workbook, builds a two‑dimensional int[,] matrix, computes its dimensions, generates a matching Range with Cells.CreateRange starting at A1, assigns the matrix to Range.Value, and saves the file as MatrixFromArray.xlsx.
public class FillMatrixFromArray
{
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array source (matrix data)
        int[,] sourceMatrix = new int[,]
        {
            { 1, 2, 3, 4 },
            { 5, 6, 7, 8 },
            { 9, 10, 11, 12 }
        };

        // Determine the size of the matrix
        int rowCount = sourceMatrix.GetLength(0);
        int columnCount = sourceMatrix.GetLength(1);

        // Create a range that matches the matrix size, starting at cell A1 (row 0, column 0)
        AsposeRange targetRange = cells.CreateRange(0, 0, rowCount, columnCount);

        // Fill the range with the two‑dimensional array using the Range.Value property
        targetRange.Value = sourceMatrix;

        // Save the workbook to a file
        string outputPath = "MatrixFromArray.xlsx";
        workbook.Save(outputPath);
    }
}
