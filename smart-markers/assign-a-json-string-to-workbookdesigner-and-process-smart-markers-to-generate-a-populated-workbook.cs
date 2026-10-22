// Title: Populate Excel smart markers from a JSON string using Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Load a JSON document, convert the employee list into a DataSet, assign it to a WorkbookDesigner, and call Process to replace the smart marker &=Employees.Name in the worksheet. | Create or load an Excel workbook, set a DataSet with Name, Age, and Department columns as the data source for WorkbookDesigner, and generate a populated Result.xlsx file.
// Common Searches: asp.net how to bind JSON data to Aspose.Cells WorkbookDesigner smart markers | c# convert list of objects to DataSet for Aspose.Cells smart marker processing | using Aspose.Cells to fill Excel template with employee names from JSON | process smart markers in Aspose.Cells when template file is missing | populate Excel file with JSON array using WorkbookDesigner in .NET
// Tags: Aspose.Cells WorkbookDesigner JSON data source | C# convert object list to DataSet for smart markers | populate Excel smart markers from JSON with Aspose | process smart markers without existing template workbook | generate Excel file using Aspose.Cells designer and DataSet

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using System.Data;
using Aspose.Cells;

// Demonstrates how to deserialize a JSON string containing employee records, transform it into a DataSet, assign the DataSet to an Aspose.Cells WorkbookDesigner, process the smart marker &=Employees.Name in an Excel workbook (creating a simple template if none exists), and save the populated workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // JSON data for smart markers.
            string jsonData = @"{
                ""Employees"": [
                    { ""Name"": ""John Doe"", ""Age"": 30, ""Department"": ""Sales"" },
                    { ""Name"": ""Jane Smith"", ""Age"": 28, ""Department"": ""Marketing"" },
                    { ""Name"": ""Bob Johnson"", ""Age"": 35, ""Department"": ""HR"" }
                ]
            }";

            // Load template if it exists; otherwise create a simple workbook.
            const string templatePath = "Template.xlsx";
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                // Simple smart marker that will be replaced by employee names.
                sheet.Cells["A1"].PutValue("Smart Marker: &=Employees.Name");
            }

            // Deserialize JSON using System.Text.Json.
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            Root? dataSource = JsonSerializer.Deserialize<Root>(jsonData, options);
            if (dataSource == null || dataSource.Employees == null)
                throw new InvalidOperationException("Failed to deserialize JSON data.");

            // Convert the list of employees to a DataSet (required by older WorkbookDesigner API).
            DataSet dataSet = new DataSet();
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Age", typeof(int));
            employeeTable.Columns.Add("Department", typeof(string));

            foreach (Employee emp in dataSource.Employees)
            {
                employeeTable.Rows.Add(emp.Name, emp.Age, emp.Department);
            }

            dataSet.Tables.Add(employeeTable);

            // Set up the designer and process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the populated workbook.
            const string resultPath = "Result.xlsx";
            designer.Workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Classes matching the JSON structure.
    public class Root
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }
}
