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

            // Sample sensor readings matrix (rows = time points, columns = sensors)
            double[,] sensorData = new double[,]
            {
                { 12.5, 15.2, 13.8 },
                { 11.7, 14.9, 13.1 },
                { 13.0, 15.5, 14.2 },
                { 12.2, 15.0, 13.9 }
            };

            int rowCount = sensorData.GetLength(0);
            int colCount = sensorData.GetLength(1);

            // Import each row of the matrix horizontally using ImportArray
            for (int r = 0; r < rowCount; r++)
            {
                double[] rowValues = new double[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    rowValues[c] = sensorData[r, c];
                }

                // ImportArray(double[] array, int firstRow, int firstColumn, bool isVertical)
                // isVertical = false for horizontal import
                cells.ImportArray(rowValues, r, 0, false);
            }

            // Calculate column-wise averages
            double[] columnAverages = new double[colCount];
            for (int c = 0; c < colCount; c++)
            {
                double sum = 0;
                for (int r = 0; r < rowCount; r++)
                {
                    sum += sensorData[r, c];
                }
                columnAverages[c] = sum / rowCount;
            }

            // Insert a summary row after the data rows
            int summaryRowIndex = rowCount; // zero‑based index; next row after data
            // Optionally label the summary row
            cells[summaryRowIndex, 0].PutValue("Average");

            // Import the averages horizontally starting from column 1
            cells.ImportArray(columnAverages, summaryRowIndex, 1, false);

            // Save the workbook
            workbook.Save("SensorReadingsWithAverages.xlsx");
        }
    }
}