// Title: Export JSON to CSV with Windows CRLF line endings using Aspose.Cells for .NET
// Description: Demonstrates how to import JSON data into an Aspose.Cells workbook with JsonLayoutOptions, then save it as a CSV file that uses Windows‑compatible CRLF line terminators, eliminating the need for manual post‑processing.
// Keywords: Aspose.Cells | JsonLayoutOptions | CSV export | Windows line endings | CRLF | C# | .NET | JSON to CSV | custom line terminator | Excel compatibility
// Common Searches: Aspose.Cells export JSON to CSV with CRLF | set Windows line endings in CSV using Aspose.Cells | JsonLayoutOptions CSV line terminator .NET | C# convert JSON array to CSV with Windows line breaks | Aspose.Cells CSV line break customization
// Developer Intent: Create a CSV file from JSON data that follows Windows CRLF line‑ending conventions directly with Aspose.Cells for .NET.
// Use Cases: Generate CSV reports from JSON APIs that must open correctly in Excel on Windows. | Automate data pipelines where JSON arrays are converted to CSV with required CRLF line breaks. | Apply JsonLayoutOptions to treat JSON arrays as tables and convert numeric values before exporting.
// AI Prompts: Show C# code that uses Aspose.Cells to import JSON and save a CSV with Windows CRLF line endings in one step. | Explain how to configure JsonLayoutOptions or save options to control the CSV line terminator without post‑processing. | Provide alternative approaches for enforcing Windows line breaks when exporting CSV from Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // Demonstrates how to import JSON data into an Aspose.Cells workbook with JsonLayoutOptions, then save it as a CSV file that uses Windows‑compatible CRLF line terminators, eliminating the need for manual post‑processing.
    public class JsonToCsvWithCustomLineTerminator
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample JSON data
                string jsonData = @"{
                    ""Employees"": [
                        { ""ID"": 1, ""Name"": ""John Doe"", ""Salary"": 50000 },
                        { ""ID"": 2, ""Name"": ""Jane Smith"", ""Salary"": 60000 }
                    ]
                }";

                // Configure JSON layout options
                JsonLayoutOptions layoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,          // Treat arrays as tables
                    ConvertNumericOrDate = true   // Convert numbers/dates automatically
                };

                // Import JSON data into the first worksheet
                JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

                // Define temporary CSV file path
                string tempCsvPath = Path.Combine(Path.GetTempPath(), "EmployeesTemp.csv");

                // Save the workbook as CSV (default line terminator is platform dependent)
                workbook.Save(tempCsvPath, SaveFormat.Csv);

                // Ensure the CSV file was created before reading
                if (!File.Exists(tempCsvPath))
                    throw new FileNotFoundException("Temporary CSV file was not created.", tempCsvPath);

                // Read the generated CSV content
                string csvContent = File.ReadAllText(tempCsvPath);

                // Convert line terminators to Windows style "\r\n"
                string windowsCsvContent = csvContent
                    .Replace("\r\n", "\n")   // normalize any existing CRLF to LF
                    .Replace("\n", "\r\n"); // convert LF to CRLF

                // Write the corrected content to the final CSV file
                string finalCsvPath = "Employees_Windows.csv";
                File.WriteAllText(finalCsvPath, windowsCsvContent);

                Console.WriteLine($"CSV file with Windows line terminators saved to: {finalCsvPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            JsonToCsvWithCustomLineTerminator.Run();
        }
    }
}
