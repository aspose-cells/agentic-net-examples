// Title: Import JSON data into Excel using Aspose.Cells for .NET – convert JSON to a DataTable and write to a worksheet with column headers
// AI Prompts: Write C# code that parses a JSON array string, converts it to a System.Data.DataTable, and inserts the table into the first worksheet of an Aspose.Cells Workbook starting at cell A1 with headers. | Show how to use Aspose.Cells to write a DataTable to a worksheet, allowing specification of the start row, start column, and optional inclusion of column headers. | Create a C# method that saves an Aspose.Cells Workbook to a given file path, automatically creating the target directory if it does not exist. | Explain how to map JSON property names to Excel column headings when importing JSON data with Aspose.Cells.
// Common Searches: c# aspose.cells import json array to excel worksheet | convert json string to datatable and export to xlsx using aspose.cells | write datatable to excel with headers using aspose.cells c# | save aspose.cells workbook to custom folder create directory if missing | map json fields to excel columns aspose.cells c# example
// Tags: json-to-datatable conversion aspose.cells c# | write datatable to worksheet aspose.cells | export json data to xlsx aspose.cells | ensure output directory exists before saving workbook | include column headers when writing datatable aspose.cells

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonImport
{
    // The example reads a JSON array, converts it into a DataTable, creates an Aspose.Cells Workbook, writes the DataTable to the first worksheet starting at A1 with column headers, ensures the output directory exists, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Sample JSON source
                string json = @"
                [
                    { ""Name"": ""Alice"", ""Age"": 28, ""Country"": ""USA"" },
                    { ""Name"": ""Bob"",   ""Age"": 34, ""Country"": ""Canada"" },
                    { ""Name"": ""Carol"", ""Age"": 25, ""Country"": ""UK"" }
                ]";

                // Convert JSON array to a DataTable
                DataTable table = JsonToDataTable(json);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Write the DataTable to the worksheet starting at cell A1, include column headers
                WriteDataTableToSheet(sheet, table, includeHeaders: true, startRow: 0, startColumn: 0);

                // Determine output path and ensure directory exists
                string outputPath = Path.Combine(Environment.CurrentDirectory, "JsonImportResult.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to convert a JSON array to a DataTable
        private static DataTable JsonToDataTable(string json)
        {
            DataTable dt = new DataTable();

            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                if (doc.RootElement.ValueKind != JsonValueKind.Array)
                    throw new ArgumentException("JSON must be an array of objects.");

                bool columnsDefined = false;

                foreach (JsonElement element in doc.RootElement.EnumerateArray())
                {
                    if (element.ValueKind != JsonValueKind.Object)
                        continue;

                    // Define columns based on the first object
                    if (!columnsDefined)
                    {
                        foreach (JsonProperty prop in element.EnumerateObject())
                        {
                            dt.Columns.Add(prop.Name, typeof(string));
                        }
                        columnsDefined = true;
                    }

                    // Populate row values
                    DataRow row = dt.NewRow();
                    foreach (JsonProperty prop in element.EnumerateObject())
                    {
                        // Store the raw string representation of the value
                        row[prop.Name] = prop.Value.GetString() ?? prop.Value.ToString();
                    }
                    dt.Rows.Add(row);
                }
            }

            return dt;
        }

        // Writes a DataTable to a worksheet starting at the specified cell.
        private static void WriteDataTableToSheet(Worksheet sheet, DataTable table, bool includeHeaders, int startRow, int startColumn)
        {
            int currentRow = startRow;

            // Write column headers if required
            if (includeHeaders)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    sheet.Cells[currentRow, startColumn + col].PutValue(table.Columns[col].ColumnName);
                }
                currentRow++;
            }

            // Write data rows
            foreach (DataRow dr in table.Rows)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    sheet.Cells[currentRow, startColumn + col].PutValue(dr[col]);
                }
                currentRow++;
            }
        }
    }
}
