// Title: Merge multiple Excel templates and process sheet‑specific smart markers with Aspose.Cells WorkbookDesigner (C#)
// Description: C# example that loads two Excel templates (EmployeeTemplate.xlsx and ProductTemplate.xlsx), combines them into a single workbook, assigns a distinct data source to each sheet (a List<Employee> and a DataTable), processes all smart markers with WorkbookDesigner, and saves the merged result as CombinedResult.xlsx.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | merge Excel templates | multiple worksheets | set data source per sheet | C# Excel automation | combine workbooks | populate Excel with List | populate Excel with DataTable
// Common Searches: How to combine two Excel templates using Aspose.Cells | Aspose.Cells WorkbookDesigner multiple sheets different data sources | Process smart markers in merged workbook C# | Set separate data sources for each worksheet Aspose.Cells | Combine employee and product templates with smart markers
// Developer Intent: Combine several template workbooks, bind a unique data source to each worksheet, and execute all smart markers in one WorkbookDesigner session.
// Use Cases: Generate a consolidated employee‑and‑product report by merging two pre‑designed templates and filling each sheet with its own collection. | Create a multi‑sheet invoice where the staff list uses a List<Employee> and the product list uses a DataTable, processed together. | Build a dashboard that aggregates different data sets across worksheets, each driven by its own smart‑marker data source.
// AI Prompts: Show me C# code to merge three Excel templates and assign a separate data source to each sheet using Aspose.Cells WorkbookDesigner. | Provide an example of binding a DataSet with multiple tables to different worksheets' smart markers in Aspose.Cells. | Explain how to handle missing template files gracefully when combining workbooks and processing smart markers.

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMultipleTemplates
{
    // Sample data classes for two different sheets
    // C# example that loads two Excel templates (EmployeeTemplate.xlsx and ProductTemplate.xlsx), combines them into a single workbook, assigns a distinct data source to each sheet (a List<Employee> and a DataTable), processes all smart markers with WorkbookDesigner, and saves the merged result as CombinedResult.xlsx.
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
            try
            {
                // Load the first template (contains smart markers for Employee data)
                Workbook template1 = LoadWorkbook("EmployeeTemplate.xlsx");

                // Load the second template (contains smart markers for Product data)
                Workbook template2 = LoadWorkbook("ProductTemplate.xlsx");

                // Merge the second template into the first one so that both sheets exist in a single workbook
                template1.Combine(template2);

                // Create a WorkbookDesigner and assign the merged workbook to it
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = template1
                };

                // Prepare data source for the first sheet (Employees)
                List<Employee> employees = new List<Employee>
                {
                    new Employee("John Doe", 30),
                    new Employee("Jane Smith", 28)
                };
                // Bind the employee list to a data source name used in the first sheet's smart markers
                designer.SetDataSource("Employee", employees);

                // Prepare data source for the second sheet (Products)
                DataTable productTable = new DataTable("Products");
                productTable.Columns.Add("Title", typeof(string));
                productTable.Columns.Add("Price", typeof(decimal));
                productTable.Rows.Add("Laptop", 1200.00m);
                productTable.Rows.Add("Smartphone", 799.99m);
                // Bind the DataTable to a data source name used in the second sheet's smart markers
                designer.SetDataSource("Product", productTable);

                // Process all smart markers in the workbook (both sheets)
                designer.Process();

                // Save the final workbook containing data from both templates
                designer.Workbook.Save("CombinedResult.xlsx");
                Console.WriteLine("Combined workbook saved successfully as 'CombinedResult.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to load a workbook safely; creates an empty workbook if the file is missing
        private static Workbook LoadWorkbook(string filePath)
        {
            if (File.Exists(filePath))
            {
                return new Workbook(filePath);
            }
            else
            {
                Console.WriteLine($"Warning: File '{filePath}' not found. Creating an empty workbook as a placeholder.");
                Workbook wb = new Workbook();
                wb.Worksheets[0].Name = Path.GetFileNameWithoutExtension(filePath);
                return wb;
            }
        }
    }
}
