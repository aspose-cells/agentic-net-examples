// Title: Use Aspose.Cells Range to fill a date‑mapped matrix in C# (row and column headers with sequential dates)
// AI Prompts: Write a C# program that creates a workbook, populates the first row and first column with consecutive dates starting from a given start date, and writes a double[,] matrix into the intersecting cells using Aspose.Cells.Range. | Update the example to apply a numeric format to the matrix cells and automatically adjust column widths based on the date headers. | Generate code that exports the same date‑header matrix to a CSV file while preserving the first row and column as date strings.
// Common Searches: aspocells c# create range and assign double[,] matrix values | how to add sequential date headers to first row and column in Excel using Aspose.Cells | populate Excel worksheet with 5x5 numeric matrix and date headers c# | write consecutive dates to Excel cells with Aspose.Cells Range object
// Tags: Aspose.Cells CreateRange with double array | C# fill Excel date headers using Cells.PutValue | write numeric matrix to Excel range Aspose.Cells | date‑driven matrix generation in Excel C# | export Aspose.Cells workbook to CSV preserving headers

using Aspose.Cells;
using System;
using System.IO;

// Creates a new workbook, writes sequential dates to the first row and column as headers, builds a 5×5 double matrix, assigns it to the intersecting cells via Aspose.Cells.Range, and saves the file as Matrix.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the start date for mapping
            DateTime startDate = new DateTime(2023, 1, 1);

            // Define the size of the matrix (number of date rows and columns)
            int rowCount = 5; // number of dates down the rows
            int colCount = 5; // number of dates across the columns

            // -------------------------------------------------
            // Fill the first column with dates (row headers)
            // -------------------------------------------------
            for (int i = 0; i < rowCount; i++)
            {
                // Row i+1 because row 0 will hold column header dates
                sheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));
            }

            // -------------------------------------------------
            // Fill the first row with dates (column headers)
            // -------------------------------------------------
            for (int j = 0; j < colCount; j++)
            {
                // Column j+1 because column 0 holds row header dates
                sheet.Cells[0, j + 1].PutValue(startDate.AddDays(j));
            }

            // -------------------------------------------------
            // Prepare matrix data to be placed in the intersecting cells
            // -------------------------------------------------
            double[,] matrixValues = new double[rowCount, colCount];
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    // Example calculation – replace with real data as needed
                    matrixValues[i, j] = (i + 1) * (j + 1);
                }
            }

            // -------------------------------------------------
            // Use Aspose.Cells.Range to write the matrix values starting at cell (1,1)
            // -------------------------------------------------
            Aspose.Cells.Range matrixRange = sheet.Cells.CreateRange(1, 1, rowCount, colCount);
            matrixRange.Value = matrixValues;

            // -------------------------------------------------
            // Save the workbook to a file
            // -------------------------------------------------
            string outputPath = "Matrix.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
