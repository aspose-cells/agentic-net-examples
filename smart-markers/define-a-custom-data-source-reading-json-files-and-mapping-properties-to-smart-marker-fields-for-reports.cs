// Title: Generate an Excel report with Aspose.Cells smart markers by loading employee data from a JSON file in C#
// Common Searches: C# load employee list from JSON and bind to Aspose.Cells smart markers | How to use WorkbookDesigner.SetDataSource with a List<T> deserialized from JSON | Aspose.Cells smart markers example reading data from a JSON file in .NET | Populate Excel template with JSON deserialized objects using Aspose.Cells | Create employee Report.xlsx from employees.json using Aspose.Cells designer
// Tags: Aspose.Cells WorkbookDesigner SetDataSource JSON | smart markers populate Excel from C# objects | JSON deserialization List<Employee> for Excel reporting | Excel template generation using Aspose.Cells and JSON data | C# custom data source for Aspose.Cells smart markers

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Model representing a record in the JSON file
    // The sample defines an Employee model, reads employees.json into a List<Employee> using System.Text.Json, loads a workbook template that contains smart markers, assigns the list to the "Employees" smart marker via WorkbookDesigner.SetDataSource, processes the markers, and saves the populated report as Report.xlsx, with error handling for missing files and JSON parsing issues.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public double Salary { get; set; }
        public DateTime HireDate { get; set; }
    }

    // Custom data source that reads a JSON file and provides a collection of objects
    public static class JsonDataSource
    {
        // Reads the JSON file and deserializes it into a list of Employee objects
        public static List<Employee> LoadEmployees(string jsonFilePath)
        {
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"JSON file not found: {jsonFilePath}");

            string jsonContent = File.ReadAllText(jsonFilePath);
            return JsonSerializer.Deserialize<List<Employee>>(jsonContent);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the JSON data file
                string jsonPath = "employees.json";

                // Load data from JSON
                List<Employee> employees = JsonDataSource.LoadEmployees(jsonPath);

                // Path to the template workbook containing smart markers
                string templatePath = "Template.xlsx";

                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template workbook not found: {templatePath}");

                // Load the workbook that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner to work with smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set the data source for the smart marker named "Employees"
                // In the template, smart markers should be like &=Employees.Name, &.Salary, etc.
                designer.SetDataSource("Employees", employees);

                // Process the smart markers and generate the final report
                designer.Process();

                // Save the generated report
                string reportPath = "Report.xlsx";
                workbook.Save(reportPath);

                Console.WriteLine($"Report generated successfully: {reportPath}");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (JsonException ex)
            {
                Console.Error.WriteLine($"JSON parsing error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
