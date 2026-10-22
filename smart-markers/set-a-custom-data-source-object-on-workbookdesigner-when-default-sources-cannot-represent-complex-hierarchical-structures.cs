// Title: Bind a hierarchical Company‑Department‑Employee object to WorkbookDesigner using SetDataSource for Aspose.Cells smart markers (C#)
// AI Prompts: Generate C# code that loads an Excel template, creates a Company object with nested Department and Employee collections, assigns it to WorkbookDesigner via SetDataSource, processes smart markers, and saves the output file. | Show how to handle a missing template workbook by creating a fallback workbook before binding a complex data source with Aspose.Cells WorkbookDesigner.
// Common Searches: Aspose.Cells WorkbookDesigner SetDataSource hierarchical object example C# | How to bind nested collections to Excel smart markers using Aspose.Cells | C# generate Excel report from Company-Department-Employee model with Aspose.Cells | Process smart markers with custom data source when template file is absent | Using WorkbookDesigner to populate Excel template from complex object graph
// Tags: WorkbookDesigner SetDataSource hierarchical object | Aspose.Cells smart markers nested collections | C# generate Excel from company department model | fallback workbook creation Aspose.Cells | process Excel template with custom data source

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads an Excel template (or creates a new workbook if the template is missing), builds a hierarchical Company object containing Departments and Employees, sets this object as a custom data source on WorkbookDesigner via SetDataSource, processes smart markers in the template, and saves the populated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a WorkbookDesigner instance
            WorkbookDesigner designer = new WorkbookDesigner();

            // Load the template workbook (ensure template.xlsx exists in the executable directory)
            string templatePath = "template.xlsx";
            Workbook workbook;

            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                // Fallback: create a new workbook if the template is missing
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Template not found. Created new workbook.");
            }

            designer.Workbook = workbook;

            // Prepare a complex hierarchical data source
            Company dataSource = GetCompanyData();

            // Set the custom data source object on the designer (named data source)
            designer.SetDataSource("Company", dataSource);

            // Process the template with the provided data source
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Generates sample hierarchical data representing a company with departments and employees
    static Company GetCompanyData()
    {
        return new Company
        {
            Name = "Tech Corp",
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "Alice", Age = 30 },
                        new Employee { Name = "Bob", Age = 35 }
                    }
                },
                new Department
                {
                    Name = "Development",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "Charlie", Age = 28 },
                        new Employee { Name = "Diana", Age = 32 }
                    }
                }
            }
        };
    }
}

// Hierarchical data model classes
public class Company
{
    public string Name { get; set; } = string.Empty;
    public List<Department> Departments { get; set; } = new();
}

public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Employee> Employees { get; set; } = new();
}

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
