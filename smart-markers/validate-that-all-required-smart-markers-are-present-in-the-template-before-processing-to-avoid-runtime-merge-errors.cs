// Title: C# method to validate required smart markers in an Excel template with Aspose.Cells before processing
// AI Prompts: Write a C# function that loads an Excel workbook using Aspose.Cells, extracts all smart markers with WorkbookDesigner.GetSmartMarkers, compares them to a provided array of required markers, and throws an InvalidOperationException listing any missing markers. | Show how to invoke the validation function, bind a data collection to a smart‑marker name, and call WorkbookDesigner.Process only when the validation succeeds. | Create a complete console program that validates smart markers, catches missing‑marker errors, processes the workbook, and saves the result to a new file.
// Common Searches: aspnet validate smart markers in Excel template Aspose.Cells | c# check missing smart markers before WorkbookDesigner.Process | how to throw exception for absent smart markers using Aspose.Cells | verify required smart markers exist in .xlsx file with Aspose.Cells .NET
// Tags: smart marker presence validation Aspose.Cells | smart marker extraction using WorkbookDesigner | required smart markers check Excel template | invalidoperationexception for missing markers | pre‑process smart marker verification Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    // Provides a C# utility that loads an Excel template, retrieves its smart markers via WorkbookDesigner.GetSmartMarkers, compares them against a required list, and throws an InvalidOperationException with the names of any missing markers before any data binding or processing occurs.
    public static class SmartMarkerValidator
    {
        /// <summary>
        /// Checks whether all <paramref name="requiredMarkers"/> exist in the workbook template.
        /// Throws an exception if any marker is missing.
        /// </summary>
        /// <param name="templatePath">Path to the Excel template containing smart markers.</param>
        /// <param name="requiredMarkers">Array of smart marker strings that must be present.</param>
        public static void Validate(string templatePath, string[] requiredMarkers)
        {
            // Load the template workbook (using default load options)
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize the designer with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Retrieve all smart markers present in the workbook
            string[] existingMarkers = designer.GetSmartMarkers();

            // Convert to a HashSet for fast lookup (case‑insensitive)
            HashSet<string> markerSet = new HashSet<string>(existingMarkers, StringComparer.OrdinalIgnoreCase);

            // Collect missing markers for reporting
            List<string> missing = new List<string>();
            foreach (string required in requiredMarkers)
            {
                if (!markerSet.Contains(required))
                {
                    missing.Add(required);
                }
            }

            // If any required marker is absent, raise an informative exception
            if (missing.Count > 0)
            {
                string message = $"The following required smart markers are missing in the template '{Path.GetFileName(templatePath)}': {string.Join(", ", missing)}";
                throw new InvalidOperationException(message);
            }
        }
    }

    /// <summary>
    /// Demonstrates loading a template, validating smart markers, processing data, and saving the result.
    /// </summary>
    public class Program
    {
        public static void Main()
        {
            // Path to the template file that contains smart markers
            string templatePath = "TemplateWithSmartMarkers.xlsx";

            // Define the list of smart markers that must be present in the template
            string[] requiredMarkers = new[]
            {
                "&=Employees.Name",
                "&=Employees.Age",
                "&=Employees.Department"
            };

            try
            {
                // Validate the template before any processing
                SmartMarkerValidator.Validate(templatePath, requiredMarkers);
                Console.WriteLine("All required smart markers are present.");

                // Load the workbook (again) for processing
                Workbook workbook = new Workbook(templatePath);
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Sample data source
                var employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                    new Employee { Name = "Jane Smith", Age = 28, Department = "HR" }
                };

                // Bind the data source to the smart marker name "Employees"
                designer.SetDataSource("Employees", employees);

                // Process the smart markers now that we know they are all present
                designer.Process();

                // Save the processed workbook
                string outputPath = "ProcessedOutput.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle validation or processing errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simple POCO representing an employee (used as data source)
        public class Employee
        {
            public string Name { get; set; }
            public int Age { get; set; }
            public string Department { get; set; }
        }
    }
}
