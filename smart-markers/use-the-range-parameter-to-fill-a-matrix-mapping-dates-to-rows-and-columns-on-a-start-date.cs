// Title: Use Aspose.Cells Range to fill a 6×6 Excel matrix with sequential date headers and calculated values in C#
// AI Prompts: Write C# code that creates a new workbook with Aspose.Cells, builds a two‑dimensional object array where the first row and column contain consecutive dates starting from a given start date, assigns the array to a Range covering the matrix, and saves the file as an XLSX workbook. | Show how to use Cells.CreateRange and Range.Value in Aspose.Cells to populate a date‑header matrix and fill inner cells with a row‑index × column‑index calculation.
// Common Searches: aspnet how to create an Excel table with date headers using Aspose.Cells range | c# assign 2d object array to Aspose.Cells range for date matrix | aspocells fill worksheet with sequential dates in first row and column | example of using Aspose.Cells Range.Value to write a date‑header matrix to XLSX | populate Excel matrix with dates and calculated values using Aspose.Cells C#
// Tags: Aspose.Cells range value assignment | date header matrix Excel C# | populate worksheet with 2d object array | save workbook as XLSX using Aspose.Cells | matrix calculation row index column index

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDateMatrixDemo
{
    // The example creates a new workbook, constructs a 6×6 object[,] where the top row and left column contain consecutive dates starting from 2023‑01‑01, fills the remaining cells with the product of their zero‑based row and column indices, assigns the array to a Range that starts at A1, and saves the result as DateMatrix.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Cells cells = workbook.Worksheets[0].Cells;

                // Define matrix dimensions (including header row and column)
                int rowCount = 6;    // 1 header row + 5 data rows
                int columnCount = 6; // 1 header column + 5 data columns

                // Define the start date for both rows and columns
                DateTime startDate = new DateTime(2023, 1, 1);

                // Prepare a 2‑dimensional object array to hold dates and sample values
                object[,] matrix = new object[rowCount, columnCount];

                // Fill top‑left cell (optional label)
                matrix[0, 0] = "Date\\Date";

                // Fill header row with dates (columns)
                for (int col = 1; col < columnCount; col++)
                {
                    matrix[0, col] = startDate.AddDays(col - 1);
                }

                // Fill header column with dates (rows) and inner values
                for (int row = 1; row < rowCount; row++)
                {
                    // Row header date
                    matrix[row, 0] = startDate.AddDays(row - 1);

                    // Fill inner cells with a simple calculation (e.g., row index * column index)
                    for (int col = 1; col < columnCount; col++)
                    {
                        matrix[row, col] = (row - 1) * (col - 1);
                    }
                }

                // Create a range that covers the entire matrix starting at A1
                AsposeRange range = cells.CreateRange(0, 0, rowCount, columnCount);

                // Assign the 2‑D array to the range in one operation
                range.Value = matrix;

                // Save the workbook
                string outputPath = "DateMatrix.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
