// Title: Merge multiple Excel template files into a single WorkbookDesigner and assign different smart‑marker data sources per worksheet with Aspose.Cells for .NET
// AI Prompts: Load two Excel template workbooks, copy their worksheets into one WorkbookDesigner instance, bind a List<Employee> to the 'Employees' smart markers on the first sheet and a List<Product> to the 'Products' smart markers on the second sheet, then process and save the combined workbook. | Add additional template worksheets to an existing WorkbookDesigner, set unique data sources for each sheet’s smart markers, and generate a consolidated Excel file without modifying the original templates. | Create a placeholder worksheet when a template file is missing, merge it with available templates in a WorkbookDesigner, bind appropriate data collections, and produce the final output workbook.
// Common Searches: asp.net aspocells merge multiple template workbooks with smart markers | bind separate data collections to smart markers on different sheets using WorkbookDesigner | copy worksheets from other Excel files into a WorkbookDesigner workbook in C# | handle missing template files by creating placeholder sheets in Aspose.Cells | process smart markers across several worksheets in a combined workbook
// Tags: WorkbookDesigner merge template workbooks Aspose.Cells | smart markers distinct data source per sheet | add copy of worksheet to designer workbook C# | generate placeholder workbook when template not found | employee list data source for smart markers | product list data source for smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMultipleTemplates
{
    // Sample data classes for demonstration
    // The example loads two Excel template files (or creates placeholder workbooks if they are missing), initializes a WorkbookDesigner with the first template, copies additional template worksheets into the designer's workbook, assigns separate List<Employee> and List<Product> data sources to the corresponding smart‑marker names on each sheet, processes all smart markers across the combined sheets, and saves the result as CombinedOutput.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Employee(string name, int age) { Name = name; Age = age; }
    }

    public class Product
    {
        public string Title { get; set; }
        public decimal Price { get; set; }
        public Product(string title, decimal price) { Title = title; Price = price; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths to the template files (each template contains smart markers for its own data source)
            string[] templateFiles = { "TemplateEmployees.xlsx", "TemplateProducts.xlsx" };

            Workbook firstTemplate = LoadOrCreatePlaceholder(templateFiles[0], "Employees template not found");
            if (firstTemplate == null) return; // abort if loading failed unexpectedly

            // Initialize WorkbookDesigner with the first template workbook
            WorkbookDesigner designer = new WorkbookDesigner(firstTemplate);

            // Load remaining templates and copy their worksheets into the designer's workbook
            for (int i = 1; i < templateFiles.Length; i++)
            {
                Workbook temp = LoadOrCreatePlaceholder(templateFiles[i], $"Template {i + 1} not found");
                if (temp == null) continue; // skip this template but continue processing others

                // Prepare source worksheet array and destination names
                Worksheet[] sourceSheets = new Worksheet[] { temp.Worksheets[0] };
                string[] destNames = new string[] { $"SheetFromTemplate{i + 1}" };

                // Copy the worksheet from the temporary workbook into the designer's workbook
                designer.Workbook.Worksheets.AddCopy(sourceSheets, destNames);
            }

            // -----------------------------------------------------------------
            // Prepare distinct data sources for each sheet
            // -----------------------------------------------------------------

            // Data source for the first sheet (employees)
            List<Employee> employees = new List<Employee>
            {
                new Employee("John Doe", 30),
                new Employee("Jane Smith", 28)
            };
            // The smart markers in the first template should reference the name "Employees"
            designer.SetDataSource("Employees", employees);

            // Data source for the second sheet (products)
            List<Product> products = new List<Product>
            {
                new Product("Laptop", 1200.50m),
                new Product("Smartphone", 799.99m)
            };
            // The smart markers in the second template should reference the name "Products"
            designer.SetDataSource("Products", products);

            // -----------------------------------------------------------------
            // Process all smart markers across all worksheets
            // -----------------------------------------------------------------
            try
            {
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // Save the combined workbook
            // -----------------------------------------------------------------
            try
            {
                designer.Workbook.Save("CombinedOutput.xlsx");
                Console.WriteLine("Combined workbook saved as 'CombinedOutput.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }

        // Loads a workbook from file if it exists; otherwise creates a placeholder workbook.
        private static Workbook LoadOrCreatePlaceholder(string filePath, string placeholderMessage)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    return new Workbook(filePath);
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}. Creating placeholder workbook.");
                    Workbook wb = new Workbook();
                    Worksheet ws = wb.Worksheets[0];
                    ws.Name = "Placeholder";
                    ws.Cells["A1"].PutValue(placeholderMessage);
                    return wb;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load or create workbook for '{filePath}': {ex.Message}");
                return null;
            }
        }
    }
}
