// Title: Fill a 10×10 date‑mapped matrix in an Excel worksheet using Aspose.Cells Range.PutValue in C#
// Common Searches: C# Aspose.Cells how to write a 2D array to Excel starting at cell B2 | populate Excel worksheet with date headers for rows and columns using Aspose.Cells Range | create a date‑based matrix and fill numeric values with Aspose.Cells in .NET | Aspose.Cells Range.PutValue example for writing a matrix with date row and column labels | generate a 10x10 grid of dates and numbers in an Excel file using Aspose.Cells C#
// Tags: Aspose.Cells Range.PutValue 2D array | Aspose.Cells date header row column | Aspose.Cells generate date matrix C# | Aspose.Cells write numeric grid to worksheet | Aspose.Cells save workbook as .xlsx

using System;
using Aspose.Cells;

// This C# example creates a new workbook, writes a start date across the first row and down the first column to serve as headers, builds a 10 × 10 numeric matrix, writes the matrix into the sheet beginning at B2 using Range.PutValue, and saves the file as DateMatrix.xlsx while handling potential exceptions.
class DateMatrixFiller
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define start date and matrix size
            DateTime startDate = new DateTime(2023, 1, 1);
            int matrixSize = 10; // 10x10 matrix

            // Fill first row (columns) with dates, leaving column A for row dates
            for (int col = 0; col < matrixSize; col++)
            {
                sheet.Cells[0, col + 1].PutValue(startDate.AddDays(col).ToString("yyyy-MM-dd"));
            }

            // Fill first column (rows) with dates, leaving row 0 for column dates
            for (int row = 0; row < matrixSize; row++)
            {
                sheet.Cells[row + 1, 0].PutValue(startDate.AddDays(row).ToString("yyyy-MM-dd"));
            }

            // Prepare sample values for the inner matrix
            object[,] values = new object[matrixSize, matrixSize];
            for (int i = 0; i < matrixSize; i++)
            {
                for (int j = 0; j < matrixSize; j++)
                {
                    // Example value: (i+1)*(j+1)
                    values[i, j] = (i + 1) * (j + 1);
                }
            }

            // Manually write the 2‑D array into the worksheet starting at B2 (row 1, column 1)
            for (int i = 0; i < matrixSize; i++)
            {
                for (int j = 0; j < matrixSize; j++)
                {
                    sheet.Cells[i + 1, j + 1].PutValue(values[i, j]);
                }
            }

            // Save the workbook
            string outputPath = "DateMatrix.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
