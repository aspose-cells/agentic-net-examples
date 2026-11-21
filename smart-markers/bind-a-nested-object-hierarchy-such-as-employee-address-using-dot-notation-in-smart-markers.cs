// Title: Generate an Excel employee report with nested address fields using Aspose.Cells smart markers and dot notation in C#
// Common Searches: how to use dot notation in Aspose.Cells smart markers to bind nested objects in C# | C# Aspose.Cells WorkbookDesigner populate employee address hierarchy from a List | binding hierarchical data to Excel template with smart markers Aspose.Cells | create Excel report from List<Employee> with nested Address using Aspose.Cells | Aspose.Cells smart markers example for nested object binding in .NET
// Tags: Aspose.Cells WorkbookDesigner bind nested objects | smart markers nested object binding Excel C# | populate hierarchical data Excel using Aspose.Cells | C# generate employee address report with smart markers | Excel template nested object binding Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example defines Employee and Address classes, builds a simple Excel template with smart markers that use dot notation (Employees.Name, Employees.Address.City, Employees.Address.Street), binds a List<Employee> to the workbook via WorkbookDesigner.SetDataSource, processes the markers, and saves the populated report as EmployeeReport.xlsx.
public class Address
{
    public string City { get; set; } = string.Empty;
    public string Street { get; set; } = string.Empty;
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
        const string templatePath = "EmployeeTemplate.xlsx";
        const string outputPath = "EmployeeReport.xlsx";

        try
        {
            // Ensure the template file exists; create a simple one if missing.
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("&=Employees.Name");
                ws.Cells["B1"].PutValue("&=Employees.Address.City");
                ws.Cells["C1"].PutValue("&=Employees.Address.Street");
                tempWb.Save(templatePath);
            }

            // Load the template workbook that contains smart markers.
            var workbook = new Workbook(templatePath);

            // Prepare a list of employees with nested address objects.
            var employees = new List<Employee>
            {
                new Employee
                {
                    Name = "John Doe",
                    Address = new Address
                    {
                        City = "New York",
                        Street = "5th Avenue"
                    }
                },
                new Employee
                {
                    Name = "Jane Smith",
                    Address = new Address
                    {
                        City = "Los Angeles",
                        Street = "Sunset Blvd"
                    }
                }
            };

            // Bind the data source to the workbook using WorkbookDesigner (smart markers processor).
            var designer = new WorkbookDesigner(workbook);
            // The name "Employees" must match the smart marker prefix used in the template.
            designer.SetDataSource("Employees", employees);
            designer.Process(); // Process smart markers.

            // Save the populated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
