// Title: Import a double[,] sensor matrix into Excel and add a column‑average summary row using Aspose.Cells for .NET
// AI Prompts: Load a two‑dimensional double array into a worksheet horizontally with Cells.ImportArray, then compute column averages and write them to a new row. | Create a summary row labeled "Average" after the imported data and populate it with the calculated averages using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# import double[,] matrix into worksheet horizontally | how to calculate column averages and add a summary row with Aspose.Cells | using ImportArray to write sensor readings to Excel and append an average row | C# Aspose.Cells example for inserting a totals row after data import
// Tags: ImportArray matrix import Aspose.Cells | column average calculation Aspose.Cells | add totals row Excel Aspose.Cells | export sensor readings to XLSX Aspose.Cells | write double[,] to worksheet Aspose.Cells

using System;
using Aspose.Cells;

namespace SensorReadingsImport
{
    // The program creates a new workbook, imports a double[,] sensor readings matrix into the first worksheet horizontally via ImportArray, computes column‑wise averages, inserts a labeled "Average" row with those values, and saves the file as SensorReadingsWithAverages.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample sensor readings matrix (rows: samples, columns: sensors)
                double[,] readings = new double[,]
                {
                    { 12.5, 15.2, 13.8, 14.0 },
                    { 11.0, 16.5, 14.2, 13.7 },
                    { 13.3, 14.8, 15.1, 12.9 }
                };

                int rowCount = readings.GetLength(0);
                int colCount = readings.GetLength(1);

                // Import each row of the matrix horizontally using ImportArray
                for (int r = 0; r < rowCount; r++)
                {
                    double[] rowData = new double[colCount];
                    for (int c = 0; c < colCount; c++)
                    {
                        rowData[c] = readings[r, c];
                    }

                    // Import the row starting at (r, 0) horizontally (isVertical = false)
                    cells.ImportArray(rowData, r, 0, false);
                }

                // Calculate column‑wise averages
                double[] averages = new double[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    double sum = 0;
                    for (int r = 0; r < rowCount; r++)
                    {
                        sum += readings[r, c];
                    }
                    averages[c] = sum / rowCount;
                }

                // Insert a summary row after the data rows
                int summaryRowIndex = rowCount; // zero‑based index; next row after data

                // Place a label in the first column
                cells[summaryRowIndex, 0].PutValue("Average");

                // Import the averages starting from column 1
                cells.ImportArray(averages, summaryRowIndex, 1, false);

                // Save the workbook
                workbook.Save("SensorReadingsWithAverages.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
