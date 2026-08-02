// Title: Combine Multiple Excel Templates and Process Each Sheet with Separate Smart‑Marker Data Sources using Aspose.Cells WorkbookDesigner (C#)
// Description: Load a primary template into a Workbook, merge additional template workbooks, assign distinct data collections (e.g., Person, Product) to individual worksheets, process each sheet with WorkbookDesigner while preserving unknown smart markers, and save the consolidated workbook.
// Keywords: Aspose.Cells | WorkbookDesigner | C# merge Excel templates | smart markers multiple sheets | combine workbooks | set data source per worksheet | preserve unknown markers | Excel template merging example
// Common Searches: Aspose.Cells combine multiple templates | WorkbookDesigner process each worksheet separately | smart markers different data source per sheet | merge Excel files and keep unknown markers | C# load several template workbooks into one | set data source for specific sheet Aspose.Cells
// Developer Intent: Load several Excel template files into a single WorkbookDesigner, merge their sheets, and apply a unique smart‑marker data source to each worksheet.
// Use Cases: Create a combined personnel and product catalog report by merging two template files and filling each sheet with its own object list. | Generate a multi‑section financial workbook where each section uses a different template and data source, preserving any custom markers. | Automate the production of a consolidated inventory and sales dashboard by loading separate templates, merging them, and binding distinct data collections to each tab.
// AI Prompts: Write C# code that loads three Excel template files, merges them into one workbook, and processes each sheet with a different smart‑marker data source using Aspose.Cells WorkbookDesigner. | Explain how to preserve unrecognized smart markers while processing specific worksheets after combining multiple templates with Aspose.Cells. | Show how to handle missing template files gracefully when merging templates and assigning data sources per worksheet in a C# Aspose.Cells project.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMultipleTemplates
{
    // Sample data classes
    // Load a primary template into a Workbook, merge additional template workbooks, assign distinct data collections (e.g., Person, Product) to individual worksheets, process each sheet with WorkbookDesigner while preserving unknown smart markers, and save the consolidated workbook.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Person(string name, int age) { Name = name; Age = age; }
    }

    public class Product
    {
        public string Title { get; set; }
        public double Price { get; set; }
        public Product(string title, double price) { Title = title; Price = price; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Paths to template files (each template contains its own smart markers)
                string[] templateFiles = { "Template_Person.xlsx", "Template_Product.xlsx" };

                // Verify the first template exists before loading
                if (!File.Exists(templateFiles[0]))
                {
                    Console.WriteLine($"Template file not found: {templateFiles[0]}");
                    return;
                }

                // Load the first template into a workbook
                Workbook mainWorkbook = new Workbook(templateFiles[0]);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = mainWorkbook
                };

                // Load remaining templates and merge them into the main workbook
                for (int i = 1; i < templateFiles.Length; i++)
                {
                    if (!File.Exists(templateFiles[i]))
                    {
                        Console.WriteLine($"Template file not found: {templateFiles[i]}");
                        continue; // Skip missing templates
                    }

                    Workbook temp = new Workbook(templateFiles[i]);
                    // Combine adds all worksheets from the temporary workbook into the main one
                    mainWorkbook.Combine(temp);
                }

                // Prepare distinct data sources for each sheet
                List<Person> persons = new List<Person>
                {
                    new Person("John Doe", 30),
                    new Person("Jane Smith", 28)
                };

                List<Product> products = new List<Product>
                {
                    new Product("Laptop", 1200.50),
                    new Product("Smartphone", 799.99)
                };

                // Process first sheet (index 0) with Person data
                designer.SetDataSource("Person", persons);
                designer.Process(0, true); // true = preserve unrecognized markers

                // Process second sheet (index 1) with Product data
                designer.SetDataSource("Product", products);
                designer.Process(1, true);

                // Save the combined and processed workbook
                string outputPath = "CombinedOutput.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
