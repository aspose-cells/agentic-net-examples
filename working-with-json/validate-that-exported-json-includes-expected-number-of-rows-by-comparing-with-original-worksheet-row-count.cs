// Title: Check that exported JSON row count equals original Excel worksheet rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, saves it as JSON, reads the JSON file, and confirms that the JSON array length matches the worksheet's MaxDataRow + 1. | Create a .NET method that accepts a workbook path, exports the first worksheet to JSON, parses the JSON, and returns true if the row counts are identical, otherwise false.
// Common Searches: how to compare Excel worksheet row count with JSON export using Aspose.Cells C# | C# Aspose.Cells verify number of rows after saving workbook as JSON | validate data rows after converting Excel to JSON with Aspose.Cells .NET | check row count consistency between .xlsx and generated JSON file in C#
// Tags: Aspose.Cells JSON export row validation | C# check worksheet rows vs JSON array | Excel to JSON data row consistency | Aspose.Cells MaxDataRow verification

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The program loads an Excel workbook, determines the number of data rows in the first worksheet, saves the workbook as JSON, reads the generated JSON file, extracts the worksheet array, and validates that the JSON row count matches the original worksheet row count.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Ensure the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the number of rows that contain data (zero‑based + 1)
            int originalRowCount = sheet.Cells.MaxDataRow + 1;

            // Export the workbook to JSON
            string jsonPath = "output.json";
            workbook.Save(jsonPath, SaveFormat.Json);

            // Ensure the JSON file was created
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file \"{jsonPath}\" was not created.");
                return;
            }

            // Read and parse the generated JSON
            string jsonContent = File.ReadAllText(jsonPath);
            using JsonDocument doc = JsonDocument.Parse(jsonContent);
            JsonElement root = doc.RootElement;

            // Retrieve the array for the exported worksheet
            int jsonRowCount = 0;
            if (root.TryGetProperty(sheet.Name, out JsonElement rowsElement) &&
                rowsElement.ValueKind == JsonValueKind.Array)
            {
                jsonRowCount = rowsElement.GetArrayLength();
            }

            // Validate row counts
            if (originalRowCount == jsonRowCount)
            {
                Console.WriteLine($"Validation succeeded: both original and JSON contain {originalRowCount} rows.");
            }
            else
            {
                Console.WriteLine($"Validation failed: original rows = {originalRowCount}, JSON rows = {jsonRowCount}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
