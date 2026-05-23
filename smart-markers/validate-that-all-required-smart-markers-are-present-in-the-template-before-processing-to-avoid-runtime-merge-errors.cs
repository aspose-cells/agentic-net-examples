using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains smart markers
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Ensure the template file exists; if not, create a simple one with required markers
                if (!File.Exists(templatePath))
                {
                    CreateTemplateWorkbook(templatePath);
                }

                // Load the template workbook
                Workbook templateWorkbook = new Workbook(templatePath);

                // Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = templateWorkbook
                };

                // Define the list of smart markers that must be present in the template
                string[] requiredMarkers = new[]
                {
                    "&=Employees.Name",
                    "&=Employees.Age",
                    "&=Employees.Department"
                };

                // Retrieve all smart markers that actually exist in the template
                string[] existingMarkers = designer.GetSmartMarkers();

                // Determine which required markers are missing
                List<string> missingMarkers = requiredMarkers
                    .Where(r => !existingMarkers.Contains(r, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                // If any required markers are missing, abort processing and inform the user
                if (missingMarkers.Count > 0)
                {
                    Console.WriteLine("The following required smart markers are missing from the template:");
                    foreach (string marker in missingMarkers)
                    {
                        Console.WriteLine($"  {marker}");
                    }

                    throw new InvalidOperationException("Template validation failed due to missing smart markers.");
                }

                // All required markers are present – proceed with data binding and processing
                var employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                    new Employee { Name = "Jane Smith", Age = 28, Department = "HR" }
                };

                // Bind the data source to the designer
                designer.SetDataSource("Employees", employees);

                // Process the smart markers
                designer.Process();

                // Save the processed workbook
                string outputPath = "ProcessedOutput.xlsx";
                designer.Workbook.Save(outputPath);

                Console.WriteLine($"Processing completed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a minimal template workbook containing the required smart markers
        private static void CreateTemplateWorkbook(string path)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Employees";

            // Place markers in the first row
            ws.Cells["A1"].PutValue("&=Employees.Name");
            ws.Cells["B1"].PutValue("&=Employees.Age");
            ws.Cells["C1"].PutValue("&=Employees.Department");

            wb.Save(path);
        }
    }

    // Simple POCO class representing an employee record
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }
}