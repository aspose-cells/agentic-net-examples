using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class JsonToCsvWindowsLineTerminatorDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample JSON data
                string json = @"{
                    ""Products"": [
                        { ""ID"": 1, ""Name"": ""Laptop"", ""Price"": 999.99 },
                        { ""ID"": 2, ""Name"": ""Phone"",  ""Price"": 599.99 }
                    ]
                }";

                // Set JSON layout options – process arrays as tables
                JsonLayoutOptions layoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true
                };

                // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(json, worksheet.Cells, 0, 0, layoutOptions);

                // Build CSV content with Windows line terminator "\r\n"
                StringBuilder csvBuilder = new StringBuilder();

                // Determine the used range
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get the cell value as string
                        string cellText = worksheet.Cells[row, col].StringValue ?? string.Empty;

                        // Escape double quotes by doubling them
                        if (cellText.Contains("\""))
                        {
                            cellText = cellText.Replace("\"", "\"\"");
                        }

                        // Enclose in double quotes if needed
                        if (cellText.Contains(",") || cellText.Contains("\"") || cellText.Contains("\r") || cellText.Contains("\n"))
                        {
                            cellText = $"\"{cellText}\"";
                        }

                        csvBuilder.Append(cellText);

                        // Append comma if not the last column
                        if (col < maxCol)
                            csvBuilder.Append(",");
                    }

                    // Append Windows line terminator
                    csvBuilder.Append("\r\n");
                }

                // Write the CSV content to a file
                string outputPath = "JsonToCsvWindows.csv";
                File.WriteAllText(outputPath, csvBuilder.ToString(), Encoding.UTF8);

                Console.WriteLine($"CSV file saved to '{outputPath}' with Windows line terminators.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.Error.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }
}