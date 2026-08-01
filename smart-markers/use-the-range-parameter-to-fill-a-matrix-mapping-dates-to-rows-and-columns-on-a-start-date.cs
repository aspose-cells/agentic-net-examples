// Title: Aspose.Cells C# – Fill an Excel Date Matrix Using Range.CreateRange
// Description: Demonstrates how to create a new workbook, define a start date, and use Aspose.Cells Range objects to populate the first column (A2:A6) and first row (B1:F1) with sequential dates. The intersecting range (B2:F6) is then filled with calculated values, and the workbook is saved as DateMatrixOutput.xlsx.
// Keywords: Aspose.Cells C# Range | CreateRange date headers | populate Excel matrix | date matrix Excel | C# Aspose.Cells example | Excel date grid | fill cells with array | Aspose.Cells DateTime | Excel schedule matrix | Aspose.Cells sample code
// Common Searches: Aspose.Cells create range with dates | C# fill Excel row and column with sequential dates | How to set date headers in Aspose.Cells | Populate matrix in Excel using Aspose.Cells Range | Aspose.Cells date matrix example
// Developer Intent: Generate an Excel workbook where both the first row and first column contain sequential dates starting from a specified start date, and the inner cells are populated with computed data via Aspose.Cells Range objects.
// Use Cases: Build a scheduling grid that cross‑references dates on both axes for project planning. | Create a reporting sheet with date‑based row and column headers to compare metrics across time periods. | Prepare test data sets with date‑indexed rows and columns for automated validation of financial models.
// AI Prompts: Write C# code with Aspose.Cells that creates a 7×7 matrix, uses March 1, 2024 as the start date for both row and column headers, and fills the inner cells with random integers between 1 and 100. | Show how to modify the example so the start date is supplied by the user at runtime and the date headers are formatted as "yyyy‑MM‑dd". | Generate a reusable method that accepts rowCount, columnCount, and startDate parameters and returns a Workbook containing the date‑mapped matrix as demonstrated.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDateMatrixDemo
{
    // Demonstrates how to create a new workbook, define a start date, and use Aspose.Cells Range objects to populate the first column (A2:A6) and first row (B1:F1) with sequential dates. The intersecting range (B2:F6) is then filled with calculated values, and the workbook is saved as DateMatrixOutput.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define matrix dimensions and the start date
                int rowCount = 5;               // Number of data rows
                int columnCount = 5;            // Number of data columns
                DateTime startDate = new DateTime(2023, 1, 1);

                // -------------------------------------------------
                // 1. Fill the first column with dates (A2:A6)
                // -------------------------------------------------
                // Create a range that starts at row 1, column 0 (zero‑based) with rowCount rows and 1 column
                AsposeRange dateColumnRange = cells.CreateRange(firstRow: 1, firstColumn: 0, totalRows: rowCount, totalColumns: 1);

                // Prepare a 2‑D object array for the dates
                object[,] dateColumnValues = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++)
                {
                    dateColumnValues[i, 0] = startDate.AddDays(i);
                }

                // Assign the array to the range
                dateColumnRange.Value = dateColumnValues;

                // -------------------------------------------------
                // 2. Fill the first row with dates (B1:F1)
                // -------------------------------------------------
                // Create a range that starts at row 0, column 1 with 1 row and columnCount columns
                AsposeRange dateRowRange = cells.CreateRange(firstRow: 0, firstColumn: 1, totalRows: 1, totalColumns: columnCount);

                // Prepare a 2‑D object array for the dates
                object[,] dateRowValues = new object[1, columnCount];
                for (int j = 0; j < columnCount; j++)
                {
                    dateRowValues[0, j] = startDate.AddDays(j);
                }

                // Assign the array to the range
                dateRowRange.Value = dateRowValues;

                // -------------------------------------------------
                // 3. Fill the inner matrix with sample data (B2:F6)
                // -------------------------------------------------
                // Create a range that starts at row 1, column 1 with rowCount rows and columnCount columns
                AsposeRange dataRange = cells.CreateRange(firstRow: 1, firstColumn: 1, totalRows: rowCount, totalColumns: columnCount);

                // Prepare a 2‑D object array for the matrix values
                object[,] matrixValues = new object[rowCount, columnCount];
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        // Example calculation: (row index + 1) * (column index + 1)
                        matrixValues[i, j] = (i + 1) * (j + 1);
                    }
                }

                // Assign the array to the range
                dataRange.Value = matrixValues;

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("DateMatrixOutput.xlsx");
                Console.WriteLine("Workbook saved successfully as DateMatrixOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
