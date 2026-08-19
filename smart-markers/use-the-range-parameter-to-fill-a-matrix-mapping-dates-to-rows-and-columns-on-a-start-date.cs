// Title: Fill a Date Matrix in Excel using Aspose.Cells Range (C#)
// Description: Shows how to generate a 2‑D array of dates from a start date, assign it to an Aspose.Cells Range, and save the result as an Excel workbook.
// Keywords: Aspose.Cells | C# | Range | date matrix | populate Excel cells | 2D array | date grid | worksheet automation | Excel file generation
// Common Searches: Aspose.Cells fill range with dates C# | How to assign a 2D array to an Excel range using Aspose | Create a weekly date grid in a worksheet with Aspose.Cells | Populate Excel cells sequentially by row and column offsets | C# code to generate a date matrix in Excel
// Developer Intent: Create a rectangular block of cells where each cell holds a date incremented by the sum of its row and column positions.
// Use Cases: Generate a calendar view for a week or month with work‑day columns. | Build a scheduling matrix where dates shift across rows and columns. | Prepare a template for a date‑based heat map or Gantt chart.
// AI Prompts: Modify the example to use a custom start date and apply a short‑date number format to the filled cells. | Add header rows and columns for days and weeks while still using Range.Value to populate the date matrix. | Show how to apply conditional formatting to highlight weekends after the date matrix is filled.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Shows how to generate a 2‑D array of dates from a start date, assign it to an Aspose.Cells Range, and save the result as an Excel workbook.
    public class FillDateMatrixWithRange
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Define matrix dimensions and start position (zero‑based indices)
            int startRow = 0;          // A1
            int startColumn = 0;       // A1
            int totalRows = 7;         // e.g., a week
            int totalColumns = 5;      // e.g., work days

            // Define the start date for the matrix
            DateTime startDate = new DateTime(2023, 1, 1);

            // Prepare a two‑dimensional array to hold the date values
            // Each cell will contain a date offset by (row index + column index) days
            object[,] dateValues = new object[totalRows, totalColumns];
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    dateValues[i, j] = startDate.AddDays(i + j);
                }
            }

            // Create a range that covers the target matrix area
            AsposeRange matrixRange = cells.CreateRange(startRow, startColumn, totalRows, totalColumns);

            // Assign the prepared 2‑D array to the range; Aspose.Cells will populate each cell
            matrixRange.Value = dateValues;

            // Save the workbook to verify the result
            string outputPath = "DateMatrix.xlsx";
            workbook.Save(outputPath);
        }
    }
}
