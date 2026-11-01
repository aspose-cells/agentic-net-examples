// Title: Import a double‑dimensional sensor matrix into Excel and append a row of column averages using Aspose.Cells for .NET
// AI Prompts: Load a double[,] sensorReadings array into the first worksheet at cell A1 with ImportTwoDimensionArray, then iterate through each column to compute its average value. | Create a new row below the imported matrix, label it "Average", write the calculated column averages into that row, and save the workbook as an .xlsx file.
// Common Searches: how to use ImportTwoDimensionArray to load a double[,] into an Excel sheet with Aspose.Cells .NET | calculate column-wise averages after importing a data matrix in Aspose.Cells | add a summary row with average values to an Excel workbook using Aspose.Cells for C# | save a workbook after inserting computed averages with Aspose.Cells .NET
// Tags: ImportTwoDimensionArray double[,] to worksheet | column average calculation Aspose.Cells | append summary row Excel .NET | sensor readings export Aspose.Cells | save workbook as .xlsx with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace SensorDataProcessing
{
    // The example creates a Workbook, imports a double[,] sensor readings matrix into the first worksheet using ImportTwoDimensionArray, adds a labeled "Average" row beneath the data, computes each column's average, writes those values into the new row, ensures the output directory exists, and saves the file as SensorReadingsWithAverages.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Sample sensor readings matrix (rows: measurements, columns: sensors)
                double[,] sensorReadings = new double[,]
                {
                    { 23.5, 45.2, 12.8 },
                    { 24.1, 44.9, 13.0 },
                    { 22.8, 46.1, 12.5 },
                    { 23.9, 45.5, 13.2 }
                };

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Convert the double[,] to object[,] because ImportTwoDimensionArray expects an object array
                int rowCount = sensorReadings.GetLength(0);
                int colCount = sensorReadings.GetLength(1);
                object[,] sensorObj = new object[rowCount, colCount];
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        sensorObj[r, c] = sensorReadings[r, c];
                    }
                }

                // Import the sensor readings matrix starting at cell A1
                sheet.Cells.ImportTwoDimensionArray(sensorObj, 0, 0);

                // Row index for the summary (average) row (zero‑based)
                int summaryRowIndex = rowCount;

                // Insert a label for the summary row
                sheet.Cells[summaryRowIndex, 0].PutValue("Average");

                // Calculate and insert average for each column
                for (int col = 0; col < colCount; col++)
                {
                    double sum = 0.0;
                    for (int row = 0; row < rowCount; row++)
                    {
                        sum += sensorReadings[row, col];
                    }
                    double average = sum / rowCount;
                    sheet.Cells[summaryRowIndex, col].PutValue(average);
                }

                // Define output file path
                string outputPath = "SensorReadingsWithAverages.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
