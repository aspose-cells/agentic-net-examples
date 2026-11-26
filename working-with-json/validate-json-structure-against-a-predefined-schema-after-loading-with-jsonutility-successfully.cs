// Title: Validate an Employees JSON array with System.Text.Json and import it into an Excel workbook using Aspose.Cells in C#
// Common Searches: c# code to manually validate a JSON array schema before processing | how to use Aspose.Cells JsonUtility to import JSON data into a worksheet | setting up JsonLayoutOptions for date and number formatting in Aspose.Cells | import JSON to Excel while ignoring null values with Aspose.Cells | example of collecting validation error messages from System.Text.Json parsing
// Tags: json structure validation with System.Text.Json | Aspose.Cells JsonUtility import to worksheet | convert json array to excel table c# | ignore null values during json to excel conversion | custom date format json import Aspose.Cells | numeric formatting json import Aspose.Cells | employees json schema validation c#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonSchemaValidationDemo
{
    // The sample program defines an Employees JSON array, validates its structure using System.Text.Json by checking required fields (Id, Name, HireDate, Salary) and gathering any errors, then creates an in‑memory Aspose.Cells Workbook. It configures JsonLayoutOptions (ArrayAsTable, ConvertNumericOrDate, IgnoreNull, DateFormat, NumberFormat) and imports the validated JSON into the first worksheet starting at cell A1. Finally, the workbook is saved as Employees.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Prepare JSON data (could be read from a file)
            // ------------------------------------------------------------
            string jsonData = @"
            {
                ""Employees"": [
                    {
                        ""Id"": 1,
                        ""Name"": ""John Doe"",
                        ""HireDate"": ""2023-01-15"",
                        ""Salary"": 75000.50,
                        ""Department"": null
                    },
                    {
                        ""Id"": 2,
                        ""Name"": ""Jane Smith"",
                        ""HireDate"": ""2022-07-01"",
                        ""Salary"": 82000,
                        ""Department"": ""Finance""
                    }
                ]
            }";

            // ------------------------------------------------------------
            // 2. Validate JSON against a simple schema (manual validation)
            // ------------------------------------------------------------
            bool isValid;
            IList<string> validationErrors;

            try
            {
                isValid = ValidateJsonStructure(jsonData, out validationErrors);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error during validation: " + ex.Message);
                return;
            }

            if (!isValid)
            {
                Console.WriteLine("JSON validation failed. Errors:");
                foreach (var err in validationErrors)
                {
                    Console.WriteLine("- " + err);
                }

                // Stop further processing if validation fails.
                return;
            }

            Console.WriteLine("JSON validation succeeded.");

            // ------------------------------------------------------------
            // 3. Import the validated JSON into an Excel workbook using Aspose.Cells
            // ------------------------------------------------------------
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure import options
                JsonLayoutOptions importOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,          // Treat arrays as tables
                    ConvertNumericOrDate = true,  // Convert numbers and dates
                    IgnoreNull = true,            // Skip null values
                    IgnoreTitle = false,          // Keep titles (object property names)
                    DateFormat = "yyyy-MM-dd",    // Expected date format in JSON
                    NumberFormat = "#,##0.00"
                };

                // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(jsonData, worksheet.Cells, 0, 0, importOptions);

                // Ensure the output directory exists
                string outputPath = Path.Combine(Environment.CurrentDirectory, "Employees.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file for visual verification
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while importing JSON to Excel:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <param name="json">The JSON string to validate.</param>
        /// <param name="errors">A list that will contain validation error messages.</param>
        /// <returns>True if validation passes; otherwise false.</returns>
        private static bool ValidateJsonStructure(string json, out IList<string> errors)
        {
            errors = new List<string>();

            try
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("Employees", out JsonElement employeesElement) ||
                    employeesElement.ValueKind != JsonValueKind.Array)
                {
                    errors.Add("Root element must contain an 'Employees' array.");
                    return false;
                }

                int index = 0;
                foreach (JsonElement employee in employeesElement.EnumerateArray())
                {
                    index++;

                    // Id (integer)
                    if (!employee.TryGetProperty("Id", out JsonElement idElem) ||
                        idElem.ValueKind != JsonValueKind.Number ||
                        !idElem.TryGetInt32(out _))
                    {
                        errors.Add($"Employee #{index}: Missing or invalid 'Id'.");
                    }

                    // Name (string)
                    if (!employee.TryGetProperty("Name", out JsonElement nameElem) ||
                        nameElem.ValueKind != JsonValueKind.String)
                    {
                        errors.Add($"Employee #{index}: Missing or invalid 'Name'.");
                    }

                    // HireDate (string, date format)
                    if (!employee.TryGetProperty("HireDate", out JsonElement dateElem) ||
                        dateElem.ValueKind != JsonValueKind.String ||
                        !DateTime.TryParse(dateElem.GetString(), out _))
                    {
                        errors.Add($"Employee #{index}: Missing or invalid 'HireDate'.");
                    }

                    // Salary (number)
                    if (!employee.TryGetProperty("Salary", out JsonElement salaryElem) ||
                        salaryElem.ValueKind != JsonValueKind.Number)
                    {
                        errors.Add($"Employee #{index}: Missing or invalid 'Salary'.");
                    }
                }

                return errors.Count == 0;
            }
            catch (JsonException je)
            {
                errors.Add("Invalid JSON format: " + je.Message);
                return false;
            }
            catch (Exception ex)
            {
                errors.Add("Unexpected error during validation: " + ex.Message);
                return false;
            }
        }
    }
}
