// Title: Bind a list of Employee objects with nested Address properties to Excel smart markers using WorkbookDesigner and dot notation (C#)
// Common Searches: c# bind list of objects with nested properties to Aspose.Cells smart markers | how to use dot notation in Aspose.Cells smart markers for address fields | populate Excel template with employee address hierarchy using WorkbookDesigner | example of setting data source for complex object hierarchy in Aspose.Cells | binding nested objects to smart markers in .NET Excel automation
// Tags: WorkbookDesigner data source binding for nested objects | Aspose.Cells smart markers nested property access | Excel template population from C# object hierarchy | address property binding via dot notation in smart markers | complex type integration with Aspose.Cells smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook containing smart markers such as {Employee.Name} and {Employee.Address.Street}, creates a List<Employee> where each Employee has a nested Address object, binds the collection to the "Employee" smart marker using WorkbookDesigner.SetDataSource, processes the markers (dot notation navigates the nested Address properties), and saves the populated workbook.
public class Address
{
    public string Street { get; set; } = string.Empty;
    public string City { get; set; } = string.Empty;
}

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public Address Address { get; set; } = new Address();
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Ensure the template file exists before loading
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook that contains smart markers like {Employee.Name}, {Employee.Address.Street}, {Employee.Address.City}
            Workbook workbook = new Workbook(templatePath);

            // Prepare a collection of employees with nested address objects
            List<Employee> employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Address = new Address { Street = "123 Main St", City = "Seattle" }
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Address = new Address { Street = "456 Oak Ave", City = "Portland" }
                }
            };

            // Bind the collection to the smart marker name "Employee" using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employee", employees);
            designer.Process(); // Process smart markers; dot notation navigates nested Address object

            // Save the populated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
