// Title: Use WorkbookDesigner to populate Excel smart markers from a JSON string in C# with Aspose.Cells
// Common Searches: C# Aspose.Cells WorkbookDesigner populate smart markers from JSON string | How to convert JSON to DataSet for Aspose.Cells smart marker processing | Load Excel template with smart markers and fill it using JSON data in C# | Aspose.Cells example for processing smart markers with a JSON dataset | Populate Excel template rows using WorkbookDesigner and a JSON array in C#
// Tags: WorkbookDesigner SetDataSource DataSet example | JSON to DataSet conversion for Aspose.Cells | fill Excel smart markers via WorkbookDesigner | process smart markers in workbook template | load template.xlsx and save output.xlsx with Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for WorkbookDesigner

// The code loads an Excel template containing smart markers, converts a JSON string of employee records into a DataSet, assigns the DataSet to a WorkbookDesigner, processes the smart markers to populate the workbook, and saves the result, with error handling for missing files and invalid JSON.
class Program
{
    static void Main()
    {
        try
        {
            // JSON data that will be used to populate the smart markers
            string jsonData = @"
            {
                ""Employees"": [
                    { ""Name"": ""John Doe"", ""Age"": 30, ""Department"": ""Sales"" },
                    { ""Name"": ""Jane Smith"", ""Age"": 28, ""Department"": ""Marketing"" },
                    { ""Name"": ""Bob Johnson"", ""Age"": 35, ""Department"": ""IT"" }
                ]
            }";

            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the template workbook that contains smart markers (e.g., &=Employees.Name)
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templatePath);

            // Convert JSON to a DataSet compatible with WorkbookDesigner
            DataSet dataSet = JsonToDataSet(jsonData);

            // Assign the DataSet as the data source for the smart markers
            designer.SetDataSource(dataSet);

            // Process the smart markers and populate the workbook with data from the JSON
            designer.Process();

            // Save the populated workbook to a new file
            designer.Workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to convert the specific JSON structure to a DataSet
    private static DataSet JsonToDataSet(string json)
    {
        DataSet ds = new DataSet();
        using (JsonDocument doc = JsonDocument.Parse(json))
        {
            JsonElement root = doc.RootElement;
            if (root.TryGetProperty("Employees", out JsonElement employees))
            {
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));
                dt.Columns.Add("Department", typeof(string));

                foreach (JsonElement emp in employees.EnumerateArray())
                {
                    DataRow row = dt.NewRow();
                    row["Name"] = emp.GetProperty("Name").GetString();
                    row["Age"] = emp.GetProperty("Age").GetInt32();
                    row["Department"] = emp.GetProperty("Department").GetString();
                    dt.Rows.Add(row);
                }

                ds.Tables.Add(dt);
            }
            else
            {
                throw new InvalidOperationException("JSON does not contain 'Employees' array.");
            }
        }
        return ds;
    }
}
