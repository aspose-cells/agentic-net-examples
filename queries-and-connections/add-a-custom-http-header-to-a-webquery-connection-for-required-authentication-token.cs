// Title: How to add an Authorization Bearer header to a WebQuery request and import CSV data into an Aspose.Cells workbook (C#)
// Common Searches: C# Aspose.Cells fetch CSV from secured API using HttpClient and add Authorization header | Import data from authenticated web service into Excel worksheet with Aspose.Cells | Add custom HTTP header to WebQuery connection in Aspose.Cells C# example | Retrieve CSV with bearer token and load into Aspose.Cells workbook | Aspose.Cells WebQuery with token authentication and CSV to worksheet
// Tags: authorization header Aspose.Cells WebQuery | httpclient bearer token data retrieval C# | csv to datatable conversion Aspose.Cells | worksheet population from DataTable | save workbook after web query

using System;
using System.Data;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// The example creates an Aspose.Cells workbook, uses HttpClient with an Authorization Bearer header to download CSV data from a protected web service, converts the CSV into a DataTable, writes the table into the first worksheet starting at cell A1, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // URL of the web service providing data
            string url = "https://api.example.com/data";

            // Prepare HTTP client with authentication header
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer YOUR_AUTH_TOKEN");

                // Synchronously fetch the data (assumed to be CSV)
                HttpResponseMessage response = client.GetAsync(url).Result;
                response.EnsureSuccessStatusCode();

                string csvContent = response.Content.ReadAsStringAsync().Result ?? string.Empty;

                // Convert CSV to DataTable
                DataTable dataTable = CsvToDataTable(csvContent);

                // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
                ImportDataTableToSheet(sheet, dataTable, 0, 0);
            }

            // Save the workbook to a file
            string outputPath = "WebQueryWithHeader.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Imports a DataTable into a worksheet at the specified start row/column
    private static void ImportDataTableToSheet(Worksheet sheet, DataTable table, int startRow, int startColumn)
    {
        if (sheet == null || table == null) return;

        // Write column headers
        for (int col = 0; col < table.Columns.Count; col++)
        {
            sheet.Cells[startRow, startColumn + col].PutValue(table.Columns[col].ColumnName);
        }

        // Write data rows
        for (int row = 0; row < table.Rows.Count; row++)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                sheet.Cells[startRow + 1 + row, startColumn + col].PutValue(table.Rows[row][col]);
            }
        }
    }

    // Simple CSV parser that converts CSV text into a DataTable
    private static DataTable CsvToDataTable(string csv)
    {
        DataTable table = new DataTable();
        using (StringReader reader = new StringReader(csv))
        {
            bool isFirstLine = true;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                string[] fields = line.Split(',');

                if (isFirstLine)
                {
                    // Create columns based on the header row
                    foreach (string header in fields)
                    {
                        table.Columns.Add(header.Trim());
                    }
                    isFirstLine = false;
                }
                else
                {
                    // Add data rows
                    DataRow row = table.NewRow();
                    for (int i = 0; i < fields.Length && i < table.Columns.Count; i++)
                    {
                        row[i] = fields[i].Trim();
                    }
                    table.Rows.Add(row);
                }
            }
        }
        return table;
    }
}
