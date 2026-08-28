// Title: Fill a matrix layout in an Excel worksheet by assigning a 2D object array to an Aspose.Cells Range in C#
// AI Prompts: Assign a two‑dimensional object[,] array directly to an Aspose.Cells Range using the Range.Value property. | Create a range that matches the dimensions of a source matrix and populate it with the array data in a single step. | Generate an .xlsx file where a specified cell block is filled from a C# 2D array via Aspose.Cells.
// Common Searches: how to assign a 2d object array to an Aspose.Cells range in C# | Aspose.Cells fill Excel cells from multidimensional array | create range with same size as matrix Aspose.Cells .NET | populate Excel worksheet matrix using Range.Value property | C# example for writing object[,] to Excel with Aspose.Cells
// Tags: set Range.Value with 2d object array | create range matching matrix dimensions | populate worksheet cells via Range.Value | matrix layout fill using Aspose.Cells C# | write object[,] data to Excel range

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a 3×4 object[,] matrix, builds a matching range starting at cell A1, assigns the array directly to the range's Value property, and saves the workbook as MatrixFromArray.xlsx.
public class MatrixFiller
{
    // Fills the provided range with the values from a two‑dimensional array.
    public static void FillMatrix(AsposeRange targetRange, object[,] data)
    {
        // Directly assign the 2D array to the range's Value property.
        // The range must have the same dimensions as the array.
        targetRange.Value = data;
    }

    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example source data: a 3x4 matrix.
            object[,] sourceData = new object[,]
            {
                { "R1C1", "R1C2", "R1C3", "R1C4" },
                { "R2C1", "R2C2", "R2C3", "R2C4" },
                { "R3C1", "R3C2", "R3C3", "R3C4" }
            };

            // Determine the size of the source matrix.
            int rowCount = sourceData.GetLength(0);
            int columnCount = sourceData.GetLength(1);

            // Create a range that matches the matrix size, starting at cell A1 (row 0, column 0).
            AsposeRange matrixRange = cells.CreateRange(0, 0, rowCount, columnCount);

            // Fill the created range with the source data.
            FillMatrix(matrixRange, sourceData);

            // Save the workbook to a file.
            string outputPath = "MatrixFromArray.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        MatrixFiller.Run();
    }
}
