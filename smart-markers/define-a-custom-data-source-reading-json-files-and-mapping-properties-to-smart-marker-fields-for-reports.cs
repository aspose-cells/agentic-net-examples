// Title: Create an Excel report with Aspose.Cells smart markers by loading JSON data into a DataTable in C#
// AI Prompts: Read a JSON file into a List<Dictionary<string, object>> and transform it into a DataTable suitable for Aspose.Cells smart markers. | Set the DataTable as the 'Data' source of WorkbookDesigner, process the smart markers, and save the filled workbook as a new Excel file. | Implement robust error handling for missing JSON or template files while generating the report with Aspose.Cells.
// Common Searches: how to bind JSON data to Aspose.Cells smart markers using C# | convert List<Dictionary<string, object>> to DataTable for WorkbookDesigner | populate Excel template with smart markers from a JSON file in .NET | Aspose.Cells generate report from JSON array C# example | file not found handling for JSON or template when using Aspose.Cells smart markers
// Tags: Aspose.Cells WorkbookDesigner JSON DataTable source | C# JSON to DataTable conversion for Aspose.Cells | Excel template smart marker population with Aspose.Cells | generate report.xlsx from template.xlsx using Aspose.Cells | exception handling for missing JSON or template files in Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// // Reads a JSON array, converts it to a DataTable, assigns the table to the "Data" source of WorkbookDesigner, processes smart markers in an Excel template, and saves the populated report as report.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the JSON file containing an array of objects
            string jsonPath = "data.json";

            if (!File.Exists(jsonPath))
                throw new FileNotFoundException($"JSON data file not found: {jsonPath}");

            // Load JSON and deserialize into a list of dictionaries
            List<Dictionary<string, object>> records;
            using (StreamReader sr = new StreamReader(jsonPath))
            {
                string json = sr.ReadToEnd();
                records = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(json);
            }

            // Ensure we have a valid collection
            records ??= new List<Dictionary<string, object>>();

            // Convert the list of dictionaries to a DataTable (required by WorkbookDesigner)
            DataTable dataTable = ConvertToDataTable(records);

            // Path to the Excel template that contains smart markers (e.g., <#=Data.Name#>)
            string templatePath = "template.xlsx";

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Excel template file not found: {templatePath}");

            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);

            // Set up the WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", dataTable);
            designer.Process(); // Populate the template with JSON data

            // Save the populated report
            string outputPath = "report.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to convert a list of dictionaries to a DataTable
    private static DataTable ConvertToDataTable(List<Dictionary<string, object>> records)
    {
        DataTable table = new DataTable();

        // Determine all column names from the records
        HashSet<string> columnNames = new HashSet<string>();
        foreach (var record in records)
        {
            foreach (var key in record.Keys)
                columnNames.Add(key);
        }

        // Add columns to the DataTable
        foreach (var columnName in columnNames)
            table.Columns.Add(columnName, typeof(string));

        // Populate rows
        foreach (var record in records)
        {
            DataRow row = table.NewRow();
            foreach (var columnName in columnNames)
            {
                if (record.TryGetValue(columnName, out var value) && value != null)
                    row[columnName] = value.ToString();
                else
                    row[columnName] = DBNull.Value;
            }
            table.Rows.Add(row);
        }

        return table;
    }
}
