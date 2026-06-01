using System;
using Aspose.Cells;

namespace SensorReadingsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample sensor readings matrix (rows: time points, columns: sensors)
            double[,] sensorReadings = new double[,]
            {
                { 12.5, 15.2, 13.8 },
                { 14.1, 16.0, 15.3 },
                { 13.7, 15.5, 14.9 },
                { 15.0, 16.8, 16.2 }
            };

            int rowCount = sensorReadings.GetLength(0);
            int colCount = sensorReadings.GetLength(1);

            // Import each row of the matrix into the worksheet using ImportArray (horizontal import)
            for (int i = 0; i < rowCount; i++)
            {
                double[] rowData = new double[colCount];
                for (int j = 0; j < colCount; j++)
                {
                    rowData[j] = sensorReadings[i, j];
                }

                // Import the row horizontally starting at column 0 (A column)
                // Parameters: double[] array, firstRow (0‑based), firstColumn, isVertical = false
                cells.ImportArray(rowData, i, 0, false);
            }

            // Calculate column‑wise averages
            double[] columnAverages = new double[colCount];
            for (int j = 0; j < colCount; j++)
            {
                double sum = 0;
                for (int i = 0; i < rowCount; i++)
                {
                    sum += sensorReadings[i, j];
                }
                columnAverages[j] = sum / rowCount;
            }

            // Insert a summary row after the data rows
            int summaryRowIndex = rowCount; // 0‑based index for the new row
            // Optional label in the first column
            cells[summaryRowIndex, 0].PutValue("Average");

            // Write the average values starting from the second column
            for (int j = 1; j < colCount; j++)
            {
                cells[summaryRowIndex, j].PutValue(columnAverages[j]);
            }

            // Save the workbook to a file
            workbook.Save("SensorReadingsWithAverages.xlsx");
        }
    }
}