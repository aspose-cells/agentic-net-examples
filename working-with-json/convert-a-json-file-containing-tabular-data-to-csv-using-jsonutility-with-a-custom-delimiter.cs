// Title: Convert JSON array to CSV with a semicolon delimiter using Aspose.Cells TxtSaveOptions in C#
// Common Searches: C# Aspose.Cells convert JSON file to CSV with custom semicolon separator | How to use TxtSaveOptions to export a worksheet as CSV with UTF-8 encoding in .NET | Convert JSON array to DataTable and save as CSV using Aspose.Cells library | Aspose.Cells CSV export custom delimiter example C# | Read JSON file and write to CSV with semicolon using Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions CSV export | JSON to DataTable conversion C# | custom CSV delimiter semicolon Aspose.Cells | UTF-8 CSV output Aspose.Cells | C# workbook save as CSV

using System;
using System.Data;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Cells;

// The example reads a JSON file containing an array of objects, converts it to a DataTable, writes the data into an Aspose.Cells workbook, and saves the first worksheet as a CSV file using TxtSaveOptions with a semicolon separator and UTF-8 encoding.
class JsonToCsvConverter
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string jsonPath = "input.json";
            string csvPath = "output.csv";

            // Verify that the JSON input file exists
            if (!File.Exists(jsonPath))
            {
                Console.WriteLine($"Error: JSON file not found at '{jsonPath}'.");
                return;
            }

            // Read JSON content
            string jsonContent = File.ReadAllText(jsonPath);

            // Convert JSON to DataTable
            DataTable table = ConvertJsonToDataTable(jsonContent);
            if (table == null || table.Columns.Count == 0)
            {
                Console.WriteLine("Error: JSON content could not be converted to a DataTable.");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write column headers
            for (int col = 0; col < table.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(table.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(table.Rows[row][col]);
                }
            }

            // Configure CSV save options with a custom delimiter (semicolon)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Separator = ';',
                // Ensure UTF-8 encoding without BOM
                Encoding = Encoding.UTF8
            };

            // Save the worksheet as a CSV file using the specified options
            workbook.Save(csvPath, csvOptions);

            Console.WriteLine($"Conversion completed successfully. CSV saved to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Converts a JSON array of objects to a DataTable
    private static DataTable ConvertJsonToDataTable(string json)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            if (root.ValueKind != JsonValueKind.Array)
                return null;

            DataTable dt = new DataTable();

            bool columnsAdded = false;

            foreach (JsonElement element in root.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object)
                    continue;

                // Add columns based on the first object's property names
                if (!columnsAdded)
                {
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        dt.Columns.Add(prop.Name, typeof(string));
                    }
                    columnsAdded = true;
                }

                DataRow row = dt.NewRow();
                foreach (JsonProperty prop in element.EnumerateObject())
                {
                    row[prop.Name] = prop.Value.GetRawText().Trim('\"');
                }
                dt.Rows.Add(row);
            }

            return dt;
        }
        catch
        {
            // Return null if any parsing error occurs
            return null;
        }
    }
}
