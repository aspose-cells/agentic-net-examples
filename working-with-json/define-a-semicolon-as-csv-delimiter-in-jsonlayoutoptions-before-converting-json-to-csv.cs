using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvDelimiterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ---------- Step 1: Import JSON data using JsonLayoutOptions ----------
            // Configure JSON layout options as needed (example settings)
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                DateFormat = "yyyy-MM-dd",
                NumberFormat = "0.00"
            };

            // Sample JSON string
            string jsonData = @"{
                ""Employees"": [
                    { ""ID"": 1, ""Name"": ""John"", ""Salary"": 50000, ""HireDate"": ""2022-01-15"" },
                    { ""ID"": 2, ""Name"": ""Alice"", ""Salary"": 62000, ""HireDate"": ""2021-07-30"" }
                ]
            }";

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonData, cells, 0, 0, jsonOptions);

            // ---------- Step 2: Export the worksheet data to CSV with semicolon delimiter ----------
            // Aspose.Cells does not provide a direct CSV export option with custom delimiter,
            // so we will use the ImportCSV method as an example of setting the delimiter.
            // First, save the current worksheet to a temporary CSV file using the default comma delimiter.
            string tempCsvPath = Path.GetTempFileName();
            workbook.Save(tempCsvPath, SaveFormat.Csv);

            // Now, create a new workbook to import the CSV using a semicolon as the splitter.
            Workbook csvWorkbook = new Workbook();
            Worksheet csvWorksheet = csvWorkbook.Worksheets[0];
            Cells csvCells = csvWorksheet.Cells;

            // Import the CSV file with semicolon as the delimiter and convert numeric data.
            csvCells.ImportCSV(tempCsvPath, ";", true, 0, 0);

            // Save the final CSV file with semicolon delimiter.
            string finalCsvPath = "OutputWithSemicolonDelimiter.csv";
            csvWorkbook.Save(finalCsvPath, SaveFormat.Csv);

            // Clean up temporary file
            File.Delete(tempCsvPath);

            Console.WriteLine("CSV file saved with semicolon delimiter at: " + Path.GetFullPath(finalCsvPath));
        }
    }
}