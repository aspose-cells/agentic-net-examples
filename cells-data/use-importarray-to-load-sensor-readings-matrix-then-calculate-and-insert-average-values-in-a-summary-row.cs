// Title: Load a 2D sensor readings array into Excel and append a bold average row with Aspose.Cells for .NET
// Common Searches: how to write a double[,] matrix to an Excel worksheet using Aspose.Cells C# | calculate column-wise averages and add a summary row in an Aspose.Cells workbook | apply bold style to a specific row in an Aspose.Cells generated .xlsx file | save sensor measurement data to Excel with Aspose.Cells and include an average row | use Aspose.Cells to import array data and format a totals row in C#
// Tags: import double array to worksheet Aspose.Cells | column average calculation Aspose.Cells | format summary row bold Aspose.Cells | save workbook as xlsx Aspose.Cells | sensor data export Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a double[,] matrix of sensor readings, writes the values to a new worksheet using Aspose.Cells, computes the average for each sensor column, inserts a summary row labeled "Average" with those values, applies bold formatting to the row, and saves the result as SensorReadingsWithAverages.xlsx.
class SensorDataProcessor
{
    static void Main()
    {
        try
        {
            // Sample sensor readings matrix (rows: time points, columns: sensors)
            double[,] sensorReadings = new double[,]
            {
                { 23.5, 45.2, 12.8 },
                { 24.1, 44.9, 13.0 },
                { 22.8, 46.1, 12.5 },
                { 23.9, 45.5, 13.2 }
            };

            // Determine matrix dimensions
            int rows = sensorReadings.GetLength(0);
            int cols = sensorReadings.GetLength(1);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate the worksheet with sensor readings starting at cell A1
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(sensorReadings[r, c]);
                }
            }

            // Calculate average for each column (sensor)
            double[] columnAverages = new double[cols];
            for (int c = 0; c < cols; c++)
            {
                double sum = 0;
                for (int r = 0; r < rows; r++)
                {
                    sum += sensorReadings[r, c];
                }
                columnAverages[c] = sum / rows;
            }

            // Insert a summary row after the data (row index = rows)
            cells[rows, 0].PutValue("Average");
            for (int c = 0; c < cols; c++)
            {
                cells[rows, c].PutValue(columnAverages[c]);
            }

            // Format the summary row (bold font)
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            StyleFlag flag = new StyleFlag { FontBold = true };
            for (int c = 0; c < cols; c++)
            {
                cells[rows, c].SetStyle(style, flag);
            }

            // Define output file path
            string outputPath = "SensorReadingsWithAverages.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
